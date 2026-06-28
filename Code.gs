// ============================================================================
// SENIOR ADVIES RAPPORT ENGINE — CODE.GS V2.8.32
// Versie: 2.8.32 — 26 juni 2026
// Wijziging: PrintenBind-integratie toegevoegd. Na het opmaken van de PDF
//            wordt automatisch een druk-opdracht (magazine, A4, vouwen &
//            nieten) naar PrintenBind gestuurd. PDF wordt door PrintenBind
//            opgehaald via publieke pdf_url uit PythonAnywhere.
//            VEILIGHEIDSREM: PRINTENBIND_ACTIEF = false (niets wordt besteld
//            tot deze bewust op true staat en account is goedgekeurd).
//            maakPdfViaPythonAnywhere_ geeft nu ook pdf_url terug.
// Eerdere wijziging (2.8.31): Betalingsmail via Stripe (4 uur na rapport,
//            alleen bij voordeel >= 750). betaling_nieuw toont EUR 0 i.p.v.
//            streepje als bijdrage_jaar1_na_advies = 0. Recovery mail: geen
//            besparingssuggestie voor alleenstaanden zonder voordeel.
// ============================================================================
// CONFIG — pas hier aan, nergens anders
// ============================================================================
var SPREADSHEET_ID    = '1pUGmAWdS0YDtfJj09V8xZL2j2SUZA-Hmjpsc9QISrOE';
var CC_EMAIL          = 'intake@senior-advies.nl';
var INTERNAL_CC_EMAIL = 'intake@senior-advies.nl';
var BEDRIJF_NAAM      = 'Senior Advies';
var BEDRIJF_TEL       = '020 463 2990';
var BEDRIJF_URL       = 'senior-advies.nl';
var BEDRIJF_STAD      = 'Amstelveen';
var RAPPORT_PRIJS     = '€ 595';
var PYTHONANYWHERE_URL     = 'https://ajpunt.pythonanywhere.com/rapport';
var PYTHONANYWHERE_SLEUTEL = 'senior-advies-2026';
var STRIPE_BETAALLINK     = 'https://buy.stripe.com/7sYeVeehF9KBgso2nAa7C01';
var STRIPE_DREMPEL        = 750;   // minimaal voordeel voor betalingsmail (euro)
var BETAALMAIL_VERTRAGING = 0.017;   // TIJDELIJK 1 minuut voor testen — terugzetten naar 4
var TAB_RAPPORTAANVRAGEN  = 'Rapportaanvragen';
var TAB_BEREKENING        = 'Berekening_Rapport';
var TAB_ROUTEKAARTEN      = 'Routekaarten';
var TAB_CAK_PARAMS        = 'CAK_Parameters';
var TAB_LOG               = 'CommunicatieLog';
var V2_OFFSET = 97;
var V2 = {
  rekenmodus:      4,
  bijdragevorm:    5,
  scenarioB:       13,
  voordeelBMnd:    15,
  scenarioD:       16,
  nettoEffectDMnd: 17,
  adviesD:         18,
  jaar1Basis:      19,
  jaar1NaAdvies:   20,
  voordeelJaar1:   21,
  rendabel:        22,
  conclusie:       23,
};
// ============================================================================
// HULPFUNCTIES
// ============================================================================
function s_(v) {
  if (v === null || v === undefined) return '';
  var t = String(v).trim();
  if (t === '#NAME?' || t === '#VALUE!' || t === '#REF!' ||
      t === '#DIV/0!' || t === '#N/A' || t === '#NUM!' || t === 'undefined') return '';
  return t;
}
function n_(v) { var f = parseFloat(v); return isNaN(f) ? 0 : f; }
function ja_(v) { return s_(v).toUpperCase() === 'JA'; }
function esc_(v) {
  return String(v || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#039;');
}
function euro_(bedrag) {
  if (bedrag === null || bedrag === undefined || isNaN(bedrag)) return '\u2014';
  if (bedrag <= 0) return '\u20ac\u202f0';
  return '\u20ac\u202f' + Math.round(bedrag).toLocaleString('nl-NL');
}
function kolom_(headers, naam) {
  var i = headers.indexOf(naam);
  return i === -1 ? null : i;
}
function get_(rij, headers, naam) {
  var i = kolom_(headers, naam);
  return i !== null ? s_(rij[i]) : '';
}
function getN_(rij, headers, naam) {
  var i = kolom_(headers, naam);
  return i !== null ? n_(rij[i]) : 0;
}
function v2Get_(rij, offset) {
  var i = V2_OFFSET + offset;
  return i < rij.length ? s_(rij[i]) : '';
}
function v2N_(rij, offset) { return n_(v2Get_(rij, offset)); }
function v2Ja_(rij, offset) { return ja_(v2Get_(rij, offset)); }
function logRegel_(actie, dossierId, status, details) {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(TAB_LOG);
    if (!sheet) return;
    sheet.appendRow([new Date(), dossierId || '', '', '', actie, details || '', status || '', '', '']);
  } catch(e) {}
}
function datum_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy');
}
// ============================================================================
// STAP 1: RAPPORTAANVRAAG → BEREKENING_RAPPORT
// ============================================================================
function verwerkNieuweAanvragenV2() {
  var ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheetAanvr  = ss.getSheetByName(TAB_RAPPORTAANVRAGEN);
  var sheetBer    = ss.getSheetByName(TAB_BEREKENING);
  if (!sheetAanvr) throw new Error('Tab ' + TAB_RAPPORTAANVRAGEN + ' niet gevonden.');
  if (!sheetBer)   throw new Error('Tab ' + TAB_BEREKENING + ' niet gevonden.');
  var aanvData    = sheetAanvr.getDataRange().getValues();
  var aanvHeaders = aanvData[1].map(function(h) { return s_(h); });
  var berData     = sheetBer.getDataRange().getValues();
  var berHeaders  = berData[1].map(function(h) { return s_(h); });
  var bestaandeIds = {};
  var berRapportCol = kolom_(berHeaders, 'rapportId');
  for (var b = 2; b < berData.length; b++) {
    var id = s_(berData[b][berRapportCol]);
    if (id) bestaandeIds[id] = true;
  }
  var toegevoegd = 0;
  for (var i = 2; i < aanvData.length; i++) {
    var rij = aanvData[i];
    var quickscanId = get_(rij, aanvHeaders, 'quickscanId');
    var rapportId   = get_(rij, aanvHeaders, 'rapportId') ||
      (quickscanId ? 'RAP-' + quickscanId : 'RAP-DIRECT-' + new Date().getTime());
    if (!rapportId) continue;
    var aanvraagStatus = get_(rij, aanvHeaders, 'status');
    if (aanvraagStatus !== 'rapport_aanvraag') continue;
    var rapportVrijgevenAanvr = get_(rij, aanvHeaders, 'rapportVrijgeven');
    if (rapportVrijgevenAanvr === 'VERWERKT') continue;
    if (bestaandeIds[rapportId]) continue;
    var nieuweRij = new Array(berHeaders.length).fill('');
    function zet(doel, bron) {
      var di = kolom_(berHeaders, doel);
      var bi = kolom_(aanvHeaders, bron);
      if (di !== null && bi !== null) nieuweRij[di] = rij[bi];
    }
    function zetWaarde(doel, waarde) {
      var di = kolom_(berHeaders, doel);
      if (di !== null) nieuweRij[di] = waarde;
    }
    zetWaarde('rapportId',            rapportId);
    zetWaarde('timestampBerekening',  new Date());
    zet('quickscanId',        'quickscanId');
    zet('naam',               'naam');
    zet('email',              'email');
    zet('aanhef',             'aanhef');
    zet('adres',              'adres');
    zet('postcode',           'postcode');
    zet('woonplaats',         'woonplaats');
    zet('contactpersoonNaam', 'contactpersoonNaam');
    zet('contactpersoonEmail','contactpersoonEmail');
    zet('schuldenBox3',       'schuldenBox3');
    zet('svbGeinformeerd',    'svbGeinformeerd');
    zet('partner',           'partner');
    zet('partnerThuis',      'partnerThuis');
    zet('relatievorm',       'relatievorm');
    zet('aowStatus',         'aowStatus');
    zet('duurzaamGescheiden','duurzaamGescheiden');
    zet('inkomenPeiljaar',   'inkomenPeiljaar');
    zet('inkomenActueel',    'inkomenActueel');
    zet('vermogenPeiljaar',  'vermogen');
    zet('vermogenActueel',   'vermogenActueel');
    zet('woning',            'woning');
    zet('woningStatus',      'woningStatus');
    zet('zorgvorm',          'zorgvorm');
    zet('opname',            'opname');
    zet('startdatumOpname',  'startdatumOpname');
    zet('opnameDuur',        'opnameDuur');
    zet('bijdrageTypeBekend','bijdrageTypeBekend');
    zet('huidigeBijdrage',   'bijdrage');
    var inkPeil  = n_(get_(rij, aanvHeaders, 'inkomenPeiljaar'));
    var inkAct   = n_(get_(rij, aanvHeaders, 'inkomenActueel'));
    zetWaarde('inkomenVerschil', inkPeil - inkAct);
    zetWaarde('inkomenGedaald',  inkAct > 0 && inkAct < inkPeil ? 'JA' : 'NEE');
    var vermPeil = n_(get_(rij, aanvHeaders, 'vermogen'));
    var vermAct  = n_(get_(rij, aanvHeaders, 'vermogenActueel'));
    zetWaarde('vermogenVerschil', vermPeil - vermAct);
    zetWaarde('bronRijRapportAanvragen', i + 1);
    zetWaarde('controleStatus', 'nieuw_v2');
    zetWaarde('rapportVrijgeven', 'JA');
    var nieuweRijNr = sheetBer.getLastRow() + 1;
    sheetBer.appendRow(nieuweRij);
    zetFormulesToeOpRij_(sheetBer, nieuweRijNr);
    var aanvrVrijIdx = kolom_(aanvHeaders, 'rapportVrijgeven');
    if (aanvrVrijIdx !== null) {
      sheetAanvr.getRange(i + 1, aanvrVrijIdx + 1).setValue('VERWERKT');
    }
    bestaandeIds[rapportId] = true;
    toegevoegd++;
    logRegel_('v2_aanvraag_verwerkt', rapportId, 'ok', 'Rij ' + (i+1) + ' \u2192 Berekening_Rapport');
  }
  SpreadsheetApp.flush();
  Logger.log('V2: ' + toegevoegd + ' nieuwe aanvragen verwerkt naar Berekening_Rapport.');
  return toegevoegd;
}
// ============================================================================
// HERSTEL VERWERKT STATUS
// ============================================================================
function herstelVerwerktStatus() {
  var ss         = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheetAanvr = ss.getSheetByName(TAB_RAPPORTAANVRAGEN);
  var sheetBer   = ss.getSheetByName(TAB_BEREKENING);
  var aanvData   = sheetAanvr.getDataRange().getValues();
  var aanvHeaders = aanvData[1].map(function(h) { return s_(h); });
  var berData    = sheetBer.getDataRange().getValues();
  var berHeaders = berData[1].map(function(h) { return s_(h); });
  var berRapportCol = kolom_(berHeaders, 'rapportId');
  var bestaandeIds = {};
  for (var b = 2; b < berData.length; b++) {
    var id = s_(berData[b][berRapportCol]);
    if (id) bestaandeIds[id] = true;
  }
  var vrijIdx   = kolom_(aanvHeaders, 'rapportVrijgeven');
  var statusIdx = kolom_(aanvHeaders, 'status');
  var qsIdx     = kolom_(aanvHeaders, 'quickscanId');
  var bijgewerkt = 0;
  for (var i = 2; i < aanvData.length; i++) {
    var rij = aanvData[i];
    var status = statusIdx !== null ? s_(rij[statusIdx]) : '';
    if (status !== 'rapport_aanvraag') continue;
    var huidigVrijgeven = vrijIdx !== null ? s_(rij[vrijIdx]) : '';
    if (huidigVrijgeven === 'VERWERKT') continue;
    var quickscanId = qsIdx !== null ? s_(rij[qsIdx]) : '';
    var rapportId = quickscanId ? 'RAP-' + quickscanId : '';
    if ((rapportId && bestaandeIds[rapportId]) || !quickscanId) {
      sheetAanvr.getRange(i + 1, vrijIdx + 1).setValue('VERWERKT');
      bijgewerkt++;
      Logger.log('VERWERKT gezet op rij ' + (i+1) + ' — ' + (rapportId || 'geen quickscanId'));
    }
  }
  SpreadsheetApp.flush();
  Logger.log('herstelVerwerktStatus klaar — ' + bijgewerkt + ' rijen bijgewerkt');
  return 'Klaar — ' + bijgewerkt + ' rijen op VERWERKT gezet';
}
// ============================================================================
// STAP 2: ROUTEKAARTEN OPHALEN
// ============================================================================
function laadRoutekaarten_() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(TAB_ROUTEKAARTEN);
  if (!sheet) throw new Error('Tab ' + TAB_ROUTEKAARTEN + ' niet gevonden.');
  var data  = sheet.getDataRange().getValues();
  var map   = {};
  for (var i = 1; i < data.length; i++) {
    var rij  = data[i];
    var code = s_(rij[0]).toUpperCase();
    if (!code) continue;
    map[code] = {
      label:          s_(rij[1]),
      actieKort:      s_(rij[2]),
      actieUitleg:    s_(rij[3]),
      documenten:     s_(rij[4]),
      instantie:      s_(rij[5]),
      telefoon:       s_(rij[6]),
      openingstijden: s_(rij[7]),
    };
  }
  return map;
}
// ============================================================================
// STAP 3: ACTIEVE ROUTES BEPALEN
// ============================================================================
function bepaalRoutes_(berRij, berHeaders) {
  var routes    = [];
  var scenarioB = v2Ja_(berRij, V2.scenarioB);
  var voordeelB = v2N_(berRij, V2.voordeelBMnd) * 12;
  var scenarioD = v2Ja_(berRij, V2.scenarioD);
  var nettoD    = v2N_(berRij, V2.nettoEffectDMnd) * 12;
  var inkPeil   = getN_(berRij, berHeaders, 'inkomenPeiljaar');
  var inkAct    = getN_(berRij, berHeaders, 'inkomenActueel');
  var vermPeil  = getN_(berRij, berHeaders, 'vermogenPeiljaar');
  var vermAct   = getN_(berRij, berHeaders, 'vermogenActueel');
  if (scenarioB && voordeelB > 0) routes.push({ code: 'BIJDRAGEVORM', voordeel: voordeelB });
  if (scenarioD) routes.push({ code: 'AOW', voordeel: nettoD > 0 ? nettoD : 0 });
  if (inkAct > 0 && inkAct < inkPeil) routes.push({ code: 'PEILJAAR', voordeel: (inkPeil - inkAct) * 0.10 });
  if (vermAct > 0 && vermAct < vermPeil) routes.push({ code: 'VERMOGEN', voordeel: (vermPeil - vermAct) * 0.04 * 0.10 });
  routes.sort(function(a, b) { return b.voordeel - a.voordeel; });
  return routes;
}
// ============================================================================
// AOW KAART
// ============================================================================
function bouwAowKaart_(berRij, kaarten) {
  var adviesD  = v2Get_(berRij, V2.adviesD);
  var nettoD   = v2N_(berRij, V2.nettoEffectDMnd);
  var omzetten = adviesD.toUpperCase().indexOf('JA') === 0;
  var basisKaart = kaarten['AOW'] || {};
  var documenten = basisKaart.documenten || '';
  if (omzetten) {
    return {
      label:          'AOW omzetting',
      actieKort:      'Zet de AOW om naar alleenstaanden-AOW',
      actieUitleg:    'Wij hebben berekend dat omzetting naar alleenstaanden-AOW voor u voordelig is. '
                    + 'De hogere AOW weegt op tegen de hogere eigen bijdrage. '
                    + 'Per saldo bespaart u ' + euro_(Math.abs(nettoD)) + ' per maand. '
                    + 'Neem contact op met de SVB om de omzetting door te voeren.',
      documenten:     documenten,
      instantie:      'SVB',
      telefoon:       '088 - 949 40 00',
      openingstijden: 'Maandag t/m vrijdag\n8.00 \u2013 17.30 uur',
    };
  } else {
    return {
      label:          'AOW \u2014 geen omzetting',
      actieKort:      'Zet de AOW niet om naar alleenstaanden-AOW',
      actieUitleg:    'Wij hebben berekend dat omzetting naar alleenstaanden-AOW voor u nadelig is. '
                    + 'De hogere AOW leidt tot een hogere eigen bijdrage die zwaarder weegt. '
                    + 'Ons advies: de AOW niet omzetten. '
                    + 'Heeft u van de SVB of een andere partij het advies gekregen dit wel te doen, '
                    + 'neem dan eerst contact met ons op.',
      documenten: '', instantie: '', telefoon: '', openingstijden: '',
    };
  }
}
// ============================================================================
// CAK PARAMETERS INLEZEN
// ============================================================================
function laadCakParameters_() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(TAB_CAK_PARAMS);
  if (!sheet) throw new Error('Tab CAK_Parameters niet gevonden');
  var data    = sheet.getDataRange().getValues();
  var headers = data[1].map(function(h) { return String(h).trim(); });
  var rij = null;
  var maxJaar = 0;
  for (var i = 2; i < data.length; i++) {
    var jaar = parseFloat(data[i][0]);
    if (jaar > maxJaar) { maxJaar = jaar; rij = data[i]; }
  }
  if (!rij) throw new Error('Geen CAK parameters gevonden');
  function p(naam) {
    var idx = headers.indexOf(naam);
    return idx !== -1 ? parseFloat(rij[idx]) || 0 : 0;
  }
  return {
    jaar:                          maxJaar,
    toetsbedrag_alleenstaand:      p('toetsbedrag_alleenstaand'),
    toetsbedrag_partner:           p('toetsbedrag_partner'),
    ZVW_aftrek:                    p('ZVW_aftrek'),
    zakKleed_alleenstaand:         p('zakKleed_alleenstaand'),
    zakKleed_partner:              p('zakKleed_partner'),
    aftrek_pensioen:               p('aftrek_pensioen'),
    aftrek_geen_pensioen:          p('aftrek_geen_pensioen'),
    vrijstelling_alleenstaand_AOW: p('vrijstelling_alleenstaand_AOW'),
    vrijstelling_partner_AOW:      p('vrijstelling_partner_AOW'),
    lage_bijdrage_min:             p('lage_bijdrage_min'),
    lage_bijdrage_max:             p('lage_bijdrage_max'),
    hoge_bijdrage_max:             p('hoge_bijdrage_max'),
    aow_gehuwd_maand:              p('aow_gehuwd_maand'),
    aow_alleenstaand_maand:        p('aow_alleenstaand_maand'),
    aanpassing_drempel:            p('aanpassing_drempel'),
  };
}
// ============================================================================
// STAP 4: BOUWRECORD
// ============================================================================
function bouwRecordVanBerekeningRij_(rij, headers) {
  function sv(v) {
    if (v === null || v === undefined) return '';
    var t = String(v).trim();
    if (t === '#NAME?' || t === '#VALUE!' || t === '#REF!' ||
        t === '#DIV/0!' || t === '#N/A' || t === '#NUM!') return '';
    return t;
  }
  function nv(v) { var f = parseFloat(v); return isNaN(f) ? 0 : f; }
  function get(naam) {
    var idx = headers.indexOf(naam);
    return idx !== -1 ? sv(rij[idx]) : '';
  }
  function getN(naam) {
    var idx = headers.indexOf(naam);
    return idx !== -1 ? nv(rij[idx]) : 0;
  }
  // euro() voor besparingsbedragen: streepje bij 0 is correct (geen besparing)
  function euro(bedrag) {
    if (!bedrag || bedrag <= 0) return '\u2014';
    return '\u20ac\u202f' + Math.round(bedrag).toLocaleString('nl-NL');
  }
  // FIX V2.8.29: euroNieuw() voor "wat u zou moeten betalen" — €0 wél tonen
  function euroNieuw(bedrag) {
    if (bedrag === null || bedrag === undefined || isNaN(bedrag)) return '\u2014';
    if (bedrag <= 0) return '\u20ac\u202f0';
    return '\u20ac\u202f' + Math.round(bedrag).toLocaleString('nl-NL');
  }
  var naam                = get('naam') || 'Cli\u00ebnt';
  var rapportId           = get('rapportId');
  var email               = get('email');
  var datum               = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy');
  var aanhef              = get('aanhef');
  var adres               = get('adres');
  var postcode            = get('postcode');
  var woonplaats          = get('woonplaats');
  var contactpersoonNaam  = get('contactpersoonNaam');
  var contactpersoonEmail = get('contactpersoonEmail');
  var H       = getN('inkomenPeiljaar');
  var L       = getN('vermogenPeiljaar');
  var O       = get('partnerThuis');
  var Q       = get('aowStatus');
  var AA      = get('bijdrageTypeBekend');
  var AB      = getN('huidigeBijdrage');
  var inkAct  = getN('inkomenActueel');
  var vermAct = getN('vermogenActueel');
  var partnerThuis = (O === 'Ja');
  var cak = laadCakParameters_();
  var CY = 12;
  var CT = partnerThuis ? cak.toetsbedrag_partner : cak.toetsbedrag_alleenstaand;
  var CU = cak.ZVW_aftrek;
  var CV = partnerThuis ? cak.zakKleed_partner : cak.zakKleed_alleenstaand;
  var CZ = L > 0 ? Math.max(0, (L - CT) * 0.04) : 0;
  var DC = 0;
  if (H > 0) {
    var netto;
    if (H <= 38441)       netto = H * 0.8152;
    else if (H <= 76817)  netto = 38441 * 0.8152 + (H - 38441) * 0.6303;
    else                  netto = 38441 * 0.8152 + (76817 - 38441) * 0.6303 + (H - 76817) * 0.505;
    DC = netto - CU - CV - cak.aftrek_pensioen;
  }
  var vrijstelling = partnerThuis ? cak.vrijstelling_partner_AOW : cak.vrijstelling_alleenstaand_AOW;
  var DD = DC > 0 ? Math.max(0, (DC - vrijstelling) * 0.25) : 0;
  var DA = H > 0
    ? Math.max(cak.lage_bijdrage_min, Math.min(cak.lage_bijdrage_max, (H + CZ) * 0.1 / 12))
    : cak.lage_bijdrage_min;
  var DB = H > 0 ? Math.min(cak.hoge_bijdrage_max, Math.max(0, (DC - DD + CZ) / 12)) : 0;
  var CX = partnerThuis ? 'LAAG' : (AA === 'Hoge eigen bijdrage' ? 'HOOG' : 'LAAG');
  var DE = CX === 'HOOG' ? DB : (AB > 0 ? AB : DA);
  var DF = (AA === 'Hoge eigen bijdrage' && partnerThuis) ? 'JA' : 'NEE';
  var scenarioB = (DF === 'JA');
  var DG = scenarioB ? DA : DE;
  var DH = scenarioB ? Math.max(0, DE - DG) : 0;
  var scenarioD = (Q.indexOf('Gehuwden') !== -1 && partnerThuis);
  var DJ = 0;
  if (scenarioD) {
    var aowVerschil = cak.aow_alleenstaand_maand - cak.aow_gehuwd_maand;
    var inkNaAow = H + aowVerschil * 12;
    var nettoNaAow;
    if (inkNaAow <= 38441)      nettoNaAow = inkNaAow * 0.8152;
    else if (inkNaAow <= 76817) nettoNaAow = 38441 * 0.8152 + (inkNaAow - 38441) * 0.6303;
    else                        nettoNaAow = 38441 * 0.8152 + (76817 - 38441) * 0.6303 + (inkNaAow - 76817) * 0.505;
    var dcNaAow = nettoNaAow - CU - cak.zakKleed_alleenstaand - cak.aftrek_pensioen;
    var ddNaAow = dcNaAow > 0 ? Math.max(0, (dcNaAow - cak.vrijstelling_alleenstaand_AOW) * 0.25) : 0;
    var daLaagNaAow = Math.max(cak.lage_bijdrage_min, Math.min(cak.lage_bijdrage_max, (inkNaAow + CZ) * 0.1 / 12));
    DJ = aowVerschil - Math.max(0, daLaagNaAow - DG);
  }
  var DK = scenarioD ? (DJ > 0 ? 'JA - omzetten voordelig' : 'NEE - niet omzetten') : 'N.v.t.';
  // AOW: als omzetting voordelig is, pas DG aan zodat voordeel_jaar1 correct wordt
  if (scenarioD && DJ > 0) {
    var dgAow = DG - DJ;
    if (dgAow < DG) DG = Math.max(0, dgAow);
  }
  // PEILJAAR: herbereken bijdrage op basis van actueel inkomen als dat lager is
  if (inkAct > 0 && inkAct < H) {
    var nettoAct;
    if (inkAct <= 38441)      nettoAct = inkAct * 0.8152;
    else if (inkAct <= 76817) nettoAct = 38441 * 0.8152 + (inkAct - 38441) * 0.6303;
    else                      nettoAct = 38441 * 0.8152 + (76817 - 38441) * 0.6303 + (inkAct - 76817) * 0.505;
    var dcAct = nettoAct - CU - CV - cak.aftrek_pensioen;
    var ddAct = dcAct > 0 ? Math.max(0, (dcAct - vrijstelling) * 0.25) : 0;
    var dbAct = Math.min(cak.hoge_bijdrage_max, Math.max(0, (dcAct - ddAct + CZ) / 12));
    var daAct = Math.max(cak.lage_bijdrage_min, Math.min(cak.lage_bijdrage_max, (inkAct + CZ) * 0.1 / 12));
    var dgAct = CX === 'HOOG' ? dbAct : daAct;
    if (dgAct < DG) DG = dgAct;
  }
  // VERMOGEN: herbereken bijdrage op basis van actueel vermogen als dat lager is
  if (vermAct >= 0 && vermAct < L) {
    var czAct = vermAct > 0 ? Math.max(0, (vermAct - CT) * 0.04) : 0;
    var inkHuidig = (inkAct > 0 && inkAct < H) ? inkAct : H;
    var dbVerm, daVerm;
    if (CX === 'HOOG') {
      var nettoVerm;
      if (inkHuidig <= 38441)      nettoVerm = inkHuidig * 0.8152;
      else if (inkHuidig <= 76817) nettoVerm = 38441 * 0.8152 + (inkHuidig - 38441) * 0.6303;
      else                         nettoVerm = 38441 * 0.8152 + (76817 - 38441) * 0.6303 + (inkHuidig - 76817) * 0.505;
      var dcVerm = nettoVerm - CU - CV - cak.aftrek_pensioen;
      var ddVerm = dcVerm > 0 ? Math.max(0, (dcVerm - vrijstelling) * 0.25) : 0;
      dbVerm = Math.min(cak.hoge_bijdrage_max, Math.max(0, (dcVerm - ddVerm + czAct) / 12));
      if (dbVerm < DG) DG = dbVerm;
    } else {
      daVerm = Math.max(cak.lage_bijdrage_min, Math.min(cak.lage_bijdrage_max, (inkHuidig + czAct) * 0.1 / 12));
      if (daVerm < DG) DG = daVerm;
    }
  }
  var bijdrage_jaar1_basis     = DE * CY;
  var bijdrage_jaar1_na_advies = Math.max(0, DG * CY);
  var voordeel_jaar1 = Math.max(0, bijdrage_jaar1_basis - bijdrage_jaar1_na_advies);
  var rendabel = voordeel_jaar1 >= 150;
  var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  var rkSheet = ss.getSheetByName(TAB_ROUTEKAARTEN);
  var kaarten = {};
  if (rkSheet) {
    var rkData = rkSheet.getDataRange().getValues();
    for (var k = 1; k < rkData.length; k++) {
      var code = String(rkData[k][0] || '').trim().toUpperCase();
      if (code) kaarten[code] = {
        actieKort:      String(rkData[k][2] || ''),
        actieUitleg:    String(rkData[k][3] || ''),
        documenten:     String(rkData[k][4] || ''),
        instantie:      String(rkData[k][5] || ''),
        telefoon:       String(rkData[k][6] || ''),
        openingstijden: String(rkData[k][7] || ''),
      };
    }
  }
  var routes = [];
  if (scenarioB && DH > 0) routes.push({ code: 'BIJDRAGEVORM', label: 'Bijdragevorm correctie', voordeel: DH * 12 });
  if (scenarioD && DJ > 0) routes.push({ code: 'AOW', label: 'AOW omzetting', voordeel: DJ * 12 });
  if (inkAct > 0 && inkAct < H) routes.push({ code: 'PEILJAAR', label: 'Inkomenswijziging', voordeel: (H - inkAct) * 0.10 });
  if (vermAct > 0 && vermAct < L) routes.push({ code: 'VERMOGEN', label: 'Vermogen aanpassing', voordeel: (L - vermAct) * 0.04 * 0.10 });
  routes.sort(function(a, b) { return b.voordeel - a.voordeel; });
  var actie1 = '', actie2 = '', actie3 = '';
  var bewijs = '';
  var instanties = [];
  var instGezien = {};
  routes.forEach(function(r, idx) {
    var kaart;
    if (r.code === 'AOW') {
      kaart = {
        actieKort:      DJ > 0 ? 'Zet de AOW om naar alleenstaanden-AOW' : 'AOW niet omzetten',
        actieUitleg:    DJ > 0
          ? 'Wij hebben berekend dat omzetting naar alleenstaanden-AOW voor u voordelig is. Per saldo bespaart u ' + euro(Math.abs(DJ)) + ' per maand. Neem contact op met de SVB.'
          : 'Wij hebben berekend dat omzetting naar alleenstaanden-AOW voor u nadelig is. Ons advies: de AOW niet omzetten.',
        documenten:     kaarten['AOW'] ? kaarten['AOW'].documenten : '',
        instantie:      'SVB',
        telefoon:       '088 - 949 40 00',
        openingstijden: 'Maandag t/m vrijdag\n8.00 \u2013 17.30 uur',
      };
    } else {
      kaart = kaarten[r.code];
    }
    if (!kaart) return;
    if (idx === 0) actie1 = kaart.actieKort + '\n\n' + kaart.actieUitleg;
    if (idx === 1) actie2 = kaart.actieKort + '\n\n' + kaart.actieUitleg;
    if (idx === 2) actie3 = kaart.actieKort + '\n\n' + kaart.actieUitleg;
    if (kaart.documenten) bewijs += (bewijs ? '\n\n' : '') + kaart.documenten;
    var instKey = kaart.instantie + '|' + kaart.telefoon;
    if (kaart.instantie && !instGezien[instKey]) {
      instGezien[instKey] = true;
      instanties.push(kaart.telefoon + ' (' + kaart.openingstijden.replace(/\n/g, ', ') + ')');
    }
  });
  var besparingRoutes = routes.filter(function(r) { return r.voordeel > 0; });
  var besparingTekst  = rendabel && besparingRoutes.length > 0
    ? besparingRoutes.map(function(r){ return r.label + ': ' + euro(r.voordeel); }).join('\n')
      + '\n\nTotaal besparing jaar 1: ' + euro(voordeel_jaar1)
    : 'Op basis van de aangeleverde gegevens is op dit moment geen duidelijke besparing vastgesteld.';
  var adresRegel = [adres, postcode + ' ' + woonplaats].filter(Boolean).join(', ');
  var instantieCak = '';
  var instantieSvb = '';
  routes.forEach(function(r) {
    var kaart = (r.code === 'AOW') ? {
      instantie: 'SVB',
      telefoon: '088 - 949 40 00',
      openingstijden: 'Maandag t/m vrijdag\n8.00 \u2013 17.30 uur'
    } : kaarten[r.code];
    if (!kaart) return;
    var regstr = kaart.telefoon + ' (' + kaart.openingstijden.replace(/\n/g, ', ') + ')';
    if (kaart.instantie === 'CAK') instantieCak = regstr;
    else if (kaart.instantie === 'SVB') instantieSvb = regstr;
  });
  return {
    rapport_ref:                 rapportId,
    rapportId:                   rapportId,
    rapportReferentie:           rapportId,
    quickscanId:                 get('quickscanId'),
    rapportDatum:                datum,
    rapport_datum:               datum,
    naam:                        naam,
    client_name:                 naam,
    clientNaam:                  naam,
    clientVolledigeNaam:         naam,
    aanhef:                      aanhef,
    aanhefNaam:                  (aanhef ? aanhef + ' ' : '') + naam,
    mailAan:                     email,
    adres:                       adres,
    postcode:                    postcode,
    woonplaats:                  woonplaats,
    adresRegel:                  adresRegel,
    contactNaam:                 contactpersoonNaam || BEDRIJF_NAAM,
    contactEmail:                contactpersoonEmail || CC_EMAIL,
    contactTelefoon:             BEDRIJF_TEL,
    besparing_jaar1:             euro(voordeel_jaar1),
    voordeel_jaar1:              voordeel_jaar1,
    bijdrage_jaar1_basis:        bijdrage_jaar1_basis,
    bijdrage_jaar1_na_advies:    bijdrage_jaar1_na_advies,
    betaling_huidig:             euro(bijdrage_jaar1_basis),
    // FIX V2.8.29: euroNieuw() zodat €0 correct wordt getoond i.p.v. streepje
    betaling_nieuw:              euroNieuw(bijdrage_jaar1_na_advies),
    rapportHoofdbevindingTekst:  rendabel
      ? 'Op basis van de aangeleverde gegevens zijn aanknopingspunten gevonden voor een verlaging van de eigen bijdrage.'
      : 'Op basis van de aangeleverde gegevens zijn op dit moment geen sterke aanknopingspunten gevonden voor verlaging van de eigen bijdrage.',
    rapportBesparingTekstNoWrap: besparingTekst,
    rapportSamenvattingKort:     besparingTekst,
    rapportAnalyseBlok1:         routes.length > 0 ? routes.map(function(r){ return r.label; }).join(' \u2014 ') : '',
    rapportAnalyseBlok2:         get('partner') === 'Ja' && get('partnerThuis') === 'Ja'
                                   ? 'Uw partner woont zelfstandig thuis. Dit is relevant voor de bijdragevorm en mogelijk ook voor de AOW-situatie.'
                                   : '',
    rapportAnalyseBlok3:         '',
    toonAnalyseBlok1:            routes.length > 0 ? 'JA' : 'NEE',
    toonAnalyseBlok2:            get('partner') === 'Ja' && get('partnerThuis') === 'Ja' ? 'JA' : 'NEE',
    toonAnalyseBlok3:            'NEE',
    rapportUitvoeringsrichtingTekst: routes.length > 0
      ? (function() {
          var r = routes[0];
          if (r.code === 'AOW') return DJ > 0 ? 'Zet de AOW om naar alleenstaanden-AOW' : 'AOW niet omzetten';
          var k = kaarten[r.code];
          return k ? k.actieKort : r.label;
        })()
      : 'Geen concrete acties gevonden op basis van de aangeleverde gegevens.',
    rapportActie1:               actie1,
    rapportActie2:               actie2,
    rapportActie3:               actie3,
    toonActie1:                  actie1 ? 'JA' : 'NEE',
    toonActie2:                  actie2 ? 'JA' : 'NEE',
    toonActie3:                  actie3 ? 'JA' : 'NEE',
    rapportBewijsstukken:        bewijs,
    toonBewijsstukken:           bewijs ? 'JA' : 'NEE',
    rapportInstantie:            instanties.join('\n'),
    rapportInstantieCak:         instantieCak,
    rapportInstantieSvb:         instantieSvb,
    toonInstantie:               instanties.length > 0 ? 'JA' : 'NEE',
    rapportControleStatus:       'V2.8.32 JS-berekening',
    toonFormulier:               'NEE',
    toonBeladvies:               'NEE',
    toonBetalingslink:           'NEE',
    toonJuridischeTekst:         'NEE',
    rapportBetaalreden:          '',
    rapportBeladvies:            '',
    rapportJuridischeTekst:      '',
    slottekst: 'Dit rapport is opgesteld door ' + BEDRIJF_NAAM + ' op basis van de door u aangeleverde gegevens. '
      + 'De berekeningen zijn indicatief. De werkelijke eigen bijdrage wordt vastgesteld door het CAK. '
      + BEDRIJF_NAAM + ' aanvaardt geen aansprakelijkheid voor afwijkingen die voortvloeien uit onjuist of onvolledig aangeleverde gegevens door de aanvrager. '
      + BEDRIJF_NAAM + ', ' + BEDRIJF_STAD + ' | ' + BEDRIJF_TEL + ' | ' + BEDRIJF_URL,
    pdfBestandsnaam: 'SeniorAdvies_' + rapportId + '.pdf',
    mailOnderwerp:   'Uw beoordeling eigen bijdrage \u2014 ' + BEDRIJF_NAAM,
    adviseur: { naam: 'Arnout J. Punt', titel: 'Adviseur eigen bijdrage', plaats: BEDRIJF_STAD },
    contact:  { adres: 'Prof. J.H. Bavincklaan 2-4, 1183 AT Amstelveen', tel: BEDRIJF_TEL, email: CC_EMAIL, web: BEDRIJF_URL },
  };
}
// ============================================================================
// PDF VIA PYTHONANYWHERE (WeasyPrint)
// WIJZIGING V2.8.32: geeft nu een object terug { file, pdfUrl } i.p.v. alleen
// het Drive-bestand. pdfUrl is de publieke URL die PrintenBind gebruikt om
// de PDF op te halen. Leeg als PythonAnywhere (nog) geen pdf_url teruggeeft.
// ============================================================================
function maakPdfViaPythonAnywhere_(record) {
  var data = {};
  Object.keys(record).forEach(function(k) { data[k] = record[k]; });
  data.sleutel = PYTHONANYWHERE_SLEUTEL;
  var response = UrlFetchApp.fetch(PYTHONANYWHERE_URL, {
    method:             'post',
    contentType:        'application/json',
    payload:            JSON.stringify(data),
    muteHttpExceptions: true,
    deadline:           120,
  });
  var code = response.getResponseCode();
  var body = response.getContentText();
  if (code !== 200) throw new Error('PythonAnywhere HTTP ' + code + ': ' + body.substring(0, 300));
  var parsed;
  try { parsed = JSON.parse(body); } catch(e) { throw new Error('Ongeldige JSON: ' + body.substring(0, 200)); }
  if (!parsed.ok) throw new Error('PythonAnywhere fout: ' + (parsed.fout || '').substring(0, 300));
  var pdfBytes = Utilities.base64Decode(parsed.pdf_base64);
  var pdfBlob  = Utilities.newBlob(pdfBytes, MimeType.PDF, record.pdfBestandsnaam || 'rapport.pdf');
  var pdfFile  = DriveApp.createFile(pdfBlob);
  return { file: pdfFile, pdfUrl: parsed.pdf_url || '' };
}
// ============================================================================
// PDF VIA APPS SCRIPT (fallback)
// Geeft hetzelfde object-formaat terug { file, pdfUrl }. pdfUrl is leeg:
// de fallback-PDF staat niet op een publieke URL, dus PrintenBind slaat over.
// ============================================================================
function maakRapportOutputPdf_(html, record) {
  var fileName = safeValue_(record.pdfBestandsnaam) || 'Senior Advies rapport.pdf';
  if (fileName.toLowerCase().indexOf('.pdf') === -1) fileName += '.pdf';
  var blob = HtmlService.createHtmlOutput(html).getBlob().getAs(MimeType.PDF).setName(fileName);
  var pdfFile = DriveApp.createFile(blob);
  return { file: pdfFile, pdfUrl: '' };
}
// ============================================================================
// STAP 5: HTML RAPPORT (fallback)
// ============================================================================
function bouwRapportOutputHtml_(record) {
  function raw(value)  { return safeValue_(value); }
  function txt(value)  { return escapeHtml_(safeValue_(value)); }
  function isJa(value) { return safeValue_(value).toUpperCase() === 'JA'; }
  function formatDateValue(value) {
    if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
      return Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd-MM-yyyy');
    }
    return safeValue_(value);
  }
  function renderInfoRow(label, value) {
    if (!raw(value)) return '';
    return '<div class="sa-info-row"><span>' + txt(label) + '</span><strong>' + txt(value) + '</strong></div>';
  }
  function renderActionCard(flag, title, text) {
    if (!isJa(flag) || !raw(text)) return '';
    var regels    = raw(text).split('\n\n');
    var actieKort = regels[0] || '';
    var uitleg    = regels.slice(1).join('\n\n');
    return '<div class="sa-card sa-action-card"><div class="sa-step">' + txt(title) + '</div>'
      + '<p style="font-weight:bold;color:#004b64;margin-bottom:4mm;">' + txt(actieKort) + '</p>'
      + (uitleg ? '<p>' + txt(uitleg) + '</p>' : '') + '</div>';
  }
  function renderSupportCard(flag, title, text) {
    if (!isJa(flag) || !raw(text)) return '';
    return '<div class="sa-card sa-support-card"><h3>' + txt(title) + '</h3><p>' + txt(text) + '</p></div>';
  }
  function renderSummaryCard(title, text) {
    if (!raw(text)) return '';
    return '<div class="sa-card sa-summary-card"><h3>' + txt(title) + '</h3><p>' + txt(text) + '</p></div>';
  }
  var clientNaam        = raw(record.clientNaam) || raw(record.naam) || '';
  var aanhefNaam        = raw(record.aanhefNaam) || clientNaam || 'daar';
  var rapportDatum      = formatDateValue(record.rapportDatum) || datum_();
  var rapportReferentie = raw(record.rapportReferentie) || raw(record.rapportId) || '';
  var adres      = raw(record.adres) || '';
  var postcode   = raw(record.postcode) || '';
  var woonplaats = raw(record.woonplaats) || '';
  var css = '@page{size:A4;margin:0;}*{box-sizing:border-box;}html,body{margin:0;padding:0;background:#f4f8fa;color:#1f2933;font-family:Arial,Helvetica,sans-serif;font-size:13px;line-height:1.5;}body{-webkit-print-color-adjust:exact;print-color-adjust:exact;}.sa-page{width:210mm;min-height:297mm;margin:0 auto;padding:24mm 26mm 22mm 26mm;background:#ffffff;page-break-after:always;break-after:page;}.sa-brand{font-size:12px;font-weight:800;letter-spacing:.9px;text-transform:uppercase;color:#004b64;margin-bottom:8mm;}.sa-kicker{display:inline-block;background:#eef7fa;color:#004b64;border-radius:999px;padding:7px 13px;font-size:10.5px;font-weight:800;margin-bottom:8mm;}.sa-accent-line{width:19mm;height:1.2mm;background:#ff532f;margin:0 0 8mm 0;}h1{font-size:32px;line-height:1.1;margin:0 0 9mm 0;color:#004b64;font-weight:800;}h2{font-size:22px;margin:0 0 7mm 0;color:#004b64;font-weight:800;}h3{font-size:14px;margin:0 0 3.5mm 0;color:#004b64;font-weight:800;}p{font-size:12.3px;line-height:1.64;margin:0 0 5mm 0;color:#253746;}.sa-card{display:block;background:#ffffff;border:1px solid #dce9ee;border-radius:14px;padding:7mm;margin-bottom:6.8mm;page-break-inside:avoid;}.sa-action-card{border-left:4px solid #ff532f;}.sa-step{display:inline-block;font-size:10.5px;font-weight:700;color:#004b64;background:#eef7fa;border-radius:999px;padding:6px 12px;margin-bottom:3.5mm;}.sa-info-row{border-bottom:1px solid #dce9ee;padding:3.8mm 0;}.sa-info-row span{display:block;color:#5f7280;font-size:10.7px;font-weight:700;}.sa-info-row strong{display:block;color:#1f2933;font-size:12.4px;font-weight:800;}.sa-amount-card{margin:0 0 9.5mm 0;border:1px solid #dce9ee;border-left:3px solid #ff532f;background:#ffffff;border-radius:16px;padding:8.8mm 9.2mm;}.sa-adresblok{margin-bottom:14mm;font-size:12.7px;line-height:1.7;}';
  var html = '<!doctype html><html lang="nl"><head><meta charset="UTF-8"><style>' + css + '</style></head><body>';
  html += '<section class="sa-page" style="padding-top:25mm;">';
  html += '<div class="sa-brand">Senior Advies</div>';
  html += '<div class="sa-kicker">Persoonlijk beoordelingsrapport</div>';
  html += '<h1>Eigen bijdrage zorg helder beoordeeld</h1>';
  html += '<div class="sa-accent-line"></div>';
  html += renderInfoRow('Voor', clientNaam);
  html += renderInfoRow('Rapportdatum', rapportDatum);
  html += renderInfoRow('Referentie', rapportReferentie);
  html += '</section>';
  html += '<section class="sa-page">';
  html += '<h2>Begeleidende brief</h2><div class="sa-accent-line"></div>';
  if (adres || woonplaats) {
    html += '<div class="sa-adresblok">' + txt(clientNaam) + '<br>';
    if (adres) html += txt(adres) + '<br>';
    if (postcode || woonplaats) html += txt((postcode + ' ' + woonplaats).trim()) + '<br>';
    html += '</div>';
  }
  html += '<p>Beste ' + txt(aanhefNaam) + ',</p>';
  html += '<p>Op basis van de door u aangeleverde gegevens hebben wij uw situatie inhoudelijk beoordeeld.</p>';
  html += '<p>Met vriendelijke groet,<br><strong>Senior Advies</strong></p>';
  html += '</section>';
  html += '<section class="sa-page">';
  html += '<h2>Hoofdconclusie</h2><div class="sa-accent-line"></div>';
  html += '<div class="sa-amount-card"><p style="font-size:10.8px;font-weight:800;text-transform:uppercase;color:#004b64;margin-bottom:3mm;">Indicatieve besparing eerste jaar</p>';
  html += '<span style="font-size:22px;font-weight:700;color:#004b64;">' + txt(raw(record.rapportBesparingTekstNoWrap)) + '</span></div>';
  html += renderSummaryCard('Hoofdbevinding', record.rapportHoofdbevindingTekst);
  html += '</section>';
  html += '<section class="sa-page">';
  html += '<h2>Vervolgstappen</h2><div class="sa-accent-line"></div>';
  html += renderActionCard(record.toonActie1, 'Eerste aanbevolen stap', record.rapportActie1);
  html += renderActionCard(record.toonActie2, 'Tweede aanbevolen stap', record.rapportActie2);
  html += renderActionCard(record.toonActie3, 'Derde aanbevolen stap', record.rapportActie3);
  html += renderSupportCard(record.toonBewijsstukken, 'Benodigde bewijsstukken', record.rapportBewijsstukken);
  html += renderSupportCard(record.toonInstantie, 'Betrokken instantie', record.rapportInstantie);
  html += '</section>';
  html += '</body></html>';
  return html;
}
// ============================================================================
// EMAIL VERSTUREN
// ============================================================================
function verstuurRapportOutputMail_(record, pdfFile) {
  var mailAan           = safeValue_(record.mailAan);
  var mailOnderwerp     = safeValue_(record.mailOnderwerp) || 'Uw persoonlijke beoordelingsrapport staat klaar';
  var clientNaam        = safeValue_(record.clientNaam) || safeValue_(record.naam) || 'daar';
  var rapportReferentie = safeValue_(record.rapportReferentie) || '';
  if (!mailAan) throw new Error('mailAan ontbreekt.');
  var aanhefRaw = safeValue_(record.aanhef || '');
  var aanhefTekst = aanhefRaw.toLowerCase().indexOf('mevr') !== -1 ? 'mevrouw'
    : aanhefRaw.toLowerCase().indexOf('dhr') !== -1 ? 'heer'
    : aanhefRaw.toLowerCase().indexOf('mevrouw') !== -1 ? 'mevrouw'
    : aanhefRaw.toLowerCase().indexOf('heer') !== -1 ? 'heer'
    : '';
  var clientNaamVolledig = safeValue_(record.clientNaam || record.naam || '');
  var geachteAanhef = aanhefTekst
    ? 'Geachte ' + aanhefTekst + ' ' + clientNaamVolledig
    : 'Geachte ' + clientNaamVolledig;
  var voordeelJaar1 = parseFloat(safeValue_(record.voordeel_jaar1) || '0') || 0;
  var besparingBedrag = voordeelJaar1 > 0
    ? '\u20ac\u202f' + Math.round(voordeelJaar1).toLocaleString('nl-NL')
    : null;
  var htmlBody = '<!DOCTYPE html><html lang="nl"><head><meta charset="UTF-8"><style>'
    + 'body{margin:0;padding:0;background:#f4f8fa;font-family:Arial,Helvetica,sans-serif;color:#1f2933;}'
    + '.wrap{max-width:600px;margin:0 auto;background:#fff;}'
    + '.header{background:#004b64;padding:28px 32px;}'
    + '.header p{color:#fff;margin:0;font-size:13px;font-weight:800;letter-spacing:.8px;text-transform:uppercase;}'
    + '.body{padding:32px;}h1{color:#004b64;font-size:22px;font-weight:800;margin:0 0 16px;}p{color:#253746;font-size:14px;line-height:1.65;margin:0 0 16px;}'
    + '.bedrag{background:#fff8f5;border-left:4px solid #ff532f;padding:14px 18px;margin:20px 0;font-size:16px;font-weight:700;color:#004b64;}'
    + '.ncnp{background:#f0f7fa;border:1px solid #dce9ee;padding:14px 18px;margin:20px 0;font-size:13px;color:#253746;}'
    + '.footer{padding:24px 32px;border-top:1px solid #dce9ee;font-size:12px;color:#6b7f8c;line-height:1.6;}'
    + '</style></head><body><div class="wrap">'
    + '<div class="header"><p>Senior Advies</p></div>'
    + '<div class="body">'
    + '<h1>Uw beoordelingsrapport staat klaar</h1>'
    + '<p>' + geachteAanhef + ',</p>'
    + '<p>Wij hebben uw situatie zorgvuldig beoordeeld. Bijgevoegd vindt u uw persoonlijk beoordelingsrapport met de bevindingen en het concrete stappenplan.</p>'
    + (besparingBedrag ? '<div class="bedrag">Mogelijke besparing eerste jaar: <strong>' + besparingBedrag + '</strong></div>' : '')
    + '<p>Wij adviseren u de aanbevolen stappen uit het rapport zo snel mogelijk op te volgen. In het rapport leest u precies wat u moet doen, welke documenten u nodig heeft en bij wie u moet zijn.</p>'
    // FIX V2.8.30: no cure no pay alleen tonen als er geen voordeel is
    + (voordeelJaar1 <= 0 ? '<div class="ncnp"><strong>No cure no pay</strong> &mdash; Blijkt uit ons rapport dat er geen voordeel te behalen is? Dan volgt er geen betalingsverplichting.</div>' : '')
    + (rapportReferentie ? '<p style="font-size:12px;color:#999;">Referentie: ' + safeValue_(rapportReferentie) + '</p>' : '')
    + '<p>Heeft u vragen? Bel ons op <strong>020 463 2990</strong> of reageer op deze mail.</p>'
    + '<p>Met vriendelijke groet,<br><strong>Arnout J. Punt</strong><br>Adviseur eigen bijdrage<br>Senior Advies</p>'
    + '</div>'
    + '<div class="footer">Senior Advies | Prof. J.H. Bavincklaan 2-4, 1183 AT Amstelveen | 020 463 2990 | intake@senior-advies.nl | senior-advies.nl</div>'
    + '</div></body></html>';
  var plainBody = geachteAanhef + ',\n\nWij hebben uw situatie zorgvuldig beoordeeld. Bijgevoegd vindt u uw persoonlijk beoordelingsrapport met de bevindingen en het concrete stappenplan.\n\n' + (besparingBedrag ? 'Mogelijke besparing eerste jaar: ' + besparingBedrag + '\n\n' : '') + 'Wij adviseren u de aanbevolen stappen zo snel mogelijk op te volgen.\n\n' + (voordeelJaar1 <= 0 ? 'No cure no pay: Blijkt uit ons rapport dat er geen voordeel te behalen is? Dan volgt er geen betalingsverplichting.\n\n' : '') + (rapportReferentie ? 'Referentie: ' + rapportReferentie + '\n\n' : '') + 'Heeft u vragen? Bel ons op 020 463 2990.\n\nMet vriendelijke groet,\nArnout J. Punt\nAdviseur eigen bijdrage\nSenior Advies\nAmstelveen | senior-advies.nl';
  GmailApp.sendEmail(mailAan, mailOnderwerp, plainBody, {
    htmlBody:    htmlBody,
    attachments: [pdfFile.getBlob()],
    name:        'Senior Advies',
    from:        'intake@senior-advies.nl',
    cc:          INTERNAL_CC_EMAIL,
  });
}

// ============================================================================
// BETALINGSMAIL VIA STRIPE
// Verstuurd 4 uur na rapport mail, alleen als voordeel >= STRIPE_DREMPEL
// ============================================================================
function verstuurBetalingsMail_(record) {
  var mailAan       = safeValue_(record.mailAan);
  if (!mailAan) return;
  var voordeelJaar1 = parseFloat(safeValue_(record.voordeel_jaar1) || '0') || 0;
  if (voordeelJaar1 < STRIPE_DREMPEL) return;

  var aanhefRaw = safeValue_(record.aanhef || '');
  var aanhefTekst = aanhefRaw.toLowerCase().indexOf('mevr') !== -1 ? 'mevrouw'
    : aanhefRaw.toLowerCase().indexOf('dhr') !== -1 ? 'heer'
    : aanhefRaw.toLowerCase().indexOf('mevrouw') !== -1 ? 'mevrouw'
    : aanhefRaw.toLowerCase().indexOf('heer') !== -1 ? 'heer'
    : '';
  var clientNaamVolledig = safeValue_(record.clientNaam || record.naam || '');
  var geachteAanhef = aanhefTekst
    ? 'Geachte ' + aanhefTekst + ' ' + clientNaamVolledig
    : 'Geachte ' + clientNaamVolledig;

  var besparingBedrag = '€ ' + Math.round(voordeelJaar1).toLocaleString('nl-NL');
  var nettoBedrag     = '€ ' + Math.round(voordeelJaar1 - 595).toLocaleString('nl-NL');
  var rapportReferentie = safeValue_(record.rapportReferentie) || '';

  var htmlBody = '<!DOCTYPE html><html lang="nl"><head><meta charset="UTF-8"><style>'
    + 'body{margin:0;padding:0;background:#f4f8fa;font-family:Arial,Helvetica,sans-serif;color:#1f2933;}'
    + '.wrap{max-width:600px;margin:0 auto;background:#fff;}'
    + '.header{background:#004b64;padding:28px 32px;}'
    + '.header p{color:#fff;margin:0;font-size:13px;font-weight:800;letter-spacing:.8px;text-transform:uppercase;}'
    + '.body{padding:32px;}h1{color:#004b64;font-size:22px;font-weight:800;margin:0 0 16px;}p{color:#253746;font-size:14px;line-height:1.65;margin:0 0 16px;}'
    + '.rekening{background:#f9f8f6;border:1px solid #e0e0e0;border-radius:8px;padding:16px 20px;margin:20px 0;}'
    + '.rekening-rij{display:flex;justify-content:space-between;padding:6px 0;border-bottom:1px solid #eee;font-size:13px;color:#444;}'
    + '.rekening-rij:last-child{border-bottom:none;font-weight:700;font-size:14px;color:#004b64;padding-top:10px;}'
    + '.betaal-knop{display:block;text-align:center;margin:24px 0;}'
    + '.betaal-knop a{display:inline-block;padding:16px 40px;background:#ff532f;color:#fff!important;text-decoration:none;border-radius:32px;font-size:16px;font-weight:800;}'
    + '.footer{padding:24px 32px;border-top:1px solid #dce9ee;font-size:12px;color:#6b7f8c;line-height:1.6;}'
    + '</style></head><body><div class="wrap">'
    + '<div class="header"><p>Senior Advies</p></div>'
    + '<div class="body">'
    + '<h1>Uw factuur staat klaar</h1>'
    + '<p>' + geachteAanhef + ',</p>'
    + '<p>Hartelijk dank voor uw vertrouwen in Senior Advies. Wij hebben een concrete besparing vastgesteld op uw eigen bijdrage. Hierbij ontvangt u onze factuur.</p>'
    + '<div class="rekening">'
    + '<div class="rekening-rij"><span>Uw besparing eerste jaar</span><span>' + besparingBedrag + '</span></div>'
    + '<div class="rekening-rij"><span>Honorarium Senior Advies</span><span>€ 595</span></div>'
    + '<div class="rekening-rij"><span>Netto voordeel voor u</span><span>' + nettoBedrag + '</span></div>'
    + '</div>'
    + '<div class="betaal-knop"><a href="' + STRIPE_BETAALLINK + '">Betaal € 595 &rarr;</a></div>'
    + '<p style="font-size:12px;color:#888;">Veilig betalen via Stripe. U ontvangt direct een betalingsbevestiging.</p>'
    + (rapportReferentie ? '<p style="font-size:12px;color:#999;">Referentie: ' + safeValue_(rapportReferentie) + '</p>' : '')
    + '<p>Heeft u vragen? Bel ons op <strong>020 463 2990</strong> of reageer op deze mail.</p>'
    + '<p>Met vriendelijke groet,<br><strong>Arnout J. Punt</strong><br>Adviseur eigen bijdrage<br>Senior Advies</p>'
    + '</div>'
    + '<div class="footer">Senior Advies | Prof. J.H. Bavincklaan 2-4, 1183 AT Amstelveen | 020 463 2990 | intake@senior-advies.nl | senior-advies.nl</div>'
    + '</div></body></html>';

  var plainBody = geachteAanhef + ',\n\nHartelijk dank voor uw vertrouwen in Senior Advies.\n\n'
    + 'Uw besparing eerste jaar: ' + besparingBedrag + '\n'
    + 'Honorarium Senior Advies: \u20ac 595\n'
    + 'Netto voordeel voor u: ' + nettoBedrag + '\n\n'
    + 'Betaal via: ' + STRIPE_BETAALLINK + '\n\n'
    + (rapportReferentie ? 'Referentie: ' + rapportReferentie + '\n\n' : '')
    + 'Heeft u vragen? Bel ons op 020 463 2990.\n\n'
    + 'Met vriendelijke groet,\nArnout J. Punt\nAdviseur eigen bijdrage\nSenior Advies\nAmstelveen | senior-advies.nl';

  GmailApp.sendEmail(mailAan, 'Uw factuur — Senior Advies', plainBody, {
    htmlBody: htmlBody,
    name:     'Senior Advies',
    from:     'intake@senior-advies.nl',
    cc:       INTERNAL_CC_EMAIL,
  });
  Logger.log('Betalingsmail verzonden: ' + (record.rapportId || '') + ' naar ' + mailAan);
}

// ============================================================================
// STATUS BIJWERKEN
// ============================================================================
function updateStatus_(rijnummer, velden) {
  var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet   = ss.getSheetByName(TAB_BEREKENING);
  var headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return s_(h); });
  Object.keys(velden).forEach(function(naam) {
    var idx = kolom_(headers, naam);
    if (idx !== null) sheet.getRange(rijnummer, idx + 1).setValue(velden[naam]);
  });
  SpreadsheetApp.flush();
}
// ============================================================================
// HOOFDFUNCTIE
// WIJZIGING V2.8.32: na het opmaken van de PDF wordt automatisch de
// PrintenBind druk-opdracht verstuurd (via verstuurPrintEnBindOrder_), VOOR
// de mail. De PDF-functie geeft nu { file, pdfUrl } terug. Een drukfout
// blokkeert het rapport en de mail niet (verstuurPrintEnBindOrder_ vangt
// fouten zelf af).
// ============================================================================
function verwerkDossier_(berRij, berHeaders, rijnummer) {
  var rapportId = get_(berRij, berHeaders, 'rapportId');
  var naam      = get_(berRij, berHeaders, 'naam');
  Logger.log('V2 start: ' + rapportId + ' (' + naam + ')');
  try {
    updateStatus_(rijnummer, { controleStatus: 'v2_render_gestart', laatsteUpdate: new Date() });
    var record = bouwRecordVanBerekeningRij_(berRij, berHeaders);
    var pdf;       // het Drive-bestand
    var pdfUrl = ''; // publieke URL voor PrintenBind
    try {
      var resultaat = maakPdfViaPythonAnywhere_(record);
      pdf    = resultaat.file;
      pdfUrl = resultaat.pdfUrl;
      Logger.log('PDF via PythonAnywhere: ' + rapportId);
    } catch(pyErr) {
      Logger.log('PythonAnywhere mislukt, fallback Apps Script: ' + pyErr.message);
      var html = bouwRapportOutputHtml_(record);
      var fb = maakRapportOutputPdf_(html, record);
      pdf    = fb.file;
      pdfUrl = fb.pdfUrl;  // leeg bij fallback -> PrintenBind slaat over
    }
    updateStatus_(rijnummer, { testPdfUrl: pdf.getUrl(), testPdfStatus: 'aangemaakt', testPdfDatum: new Date() });

    // ---- PRINTENBIND: druk-opdracht direct na opmaak PDF, voor de mail ----
    // Verstuurt alleen echt als PRINTENBIND_ACTIEF = true. Faalt dit, dan
    // gaat het rapport + de mail gewoon door (geen throw).
    var pbResultaat = verstuurPrintEnBindOrder_(record, pdfUrl);
    if (pbResultaat.ok) {
      updateStatus_(rijnummer, { printenbindStatus: 'BESTELD', printenbindOrderId: pbResultaat.orderId });
    } else {
      updateStatus_(rijnummer, { printenbindStatus: pbResultaat.melding });
    }

    verstuurRapportOutputMail_(record, pdf);
    updateStatus_(rijnummer, {
      rapportVrijgeven: 'VERZONDEN', controleStatus: 'v2_verzonden',
      testPdfStatus: 'verzonden', blokkeerReden: '', laatsteUpdate: new Date()
    });
    logRegel_('v2_verzonden', rapportId, 'ok', pdf.getUrl());
    Logger.log('V2 klaar: ' + rapportId);
    return { ok: true, rapportId: rapportId, pdfUrl: pdf.getUrl() };
  } catch (err) {
    var melding = err && err.message ? err.message : String(err);
    updateStatus_(rijnummer, { controleStatus: 'v2_fout', blokkeerReden: melding, laatsteUpdate: new Date() });
    logRegel_('v2_fout', rapportId, 'fout', melding);
    Logger.log('V2 fout: ' + rapportId + ' \u2014 ' + melding);
    throw err;
  }
}
// ============================================================================
// BATCH
// ============================================================================
function verwerkVrijgegevenDossiers() {
  var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet   = ss.getSheetByName(TAB_BEREKENING);
  if (!sheet) { Logger.log('Berekening_Rapport niet gevonden'); return; }
  var data    = sheet.getDataRange().getValues();
  var headers = data[1].map(function(h) { return s_(h); });
  var resultaat = { ok: 0, fout: 0, overgeslagen: 0 };
  for (var i = 2; i < data.length; i++) {
    var rij       = data[i];
    var rapportId = get_(rij, headers, 'rapportId');
    var vrijgeven = s_(rij[kolom_(headers, 'rapportVrijgeven') || 0]).toUpperCase();
    var status    = s_(rij[kolom_(headers, 'controleStatus') || 0]);
    if (!rapportId)                { Logger.log('Skip: geen rapportId rij '+(i+1)); resultaat.overgeslagen++; continue; }
    if (vrijgeven !== 'JA')        { Logger.log('Skip: vrijgeven='+vrijgeven+' rij '+(i+1)); resultaat.overgeslagen++; continue; }
    if (status === 'v2_verzonden') { Logger.log('Skip: al verzonden rij '+(i+1)); resultaat.overgeslagen++; continue; }
    var tsIdx = headers.indexOf('timestampBerekening');
    if (tsIdx !== -1) {
      var ts = rij[tsIdx];
      if (ts instanceof Date) {
        var uurOud = (new Date() - ts) / (1000 * 60 * 60);
        if (uurOud < 0) { Logger.log('Skip: 0u check rij '+(i+1)); resultaat.overgeslagen++; continue; }
      }
    }
    try { verwerkDossier_(rij, headers, i + 1); resultaat.ok++; }
    catch(e) { resultaat.fout++; }
  }
  Logger.log('Batch klaar \u2014 ok: ' + resultaat.ok + ', fout: ' + resultaat.fout + ', overgeslagen: ' + resultaat.overgeslagen);
  return resultaat;
}
// ============================================================================
// BETALINGSMAIL BATCH
// Verstuurt betalingsmail 4 uur na rapport voor dossiers met voordeel >= 750
// ============================================================================
function verwerkBetalingsmails() {
  var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet   = ss.getSheetByName(TAB_BEREKENING);
  if (!sheet) { Logger.log('Berekening_Rapport niet gevonden'); return; }
  var data    = sheet.getDataRange().getValues();
  var headers = data[1].map(function(h) { return s_(h); });
  var verzonden = 0;
  for (var i = 2; i < data.length; i++) {
    var rij       = data[i];
    var rapportId = get_(rij, headers, 'rapportId');
    var status    = s_(rij[kolom_(headers, 'controleStatus') || 0]);
    var betaalStatus = get_(rij, headers, 'betalingsMailStatus');
    // Alleen verzonden rapporten
    if (status !== 'v2_verzonden') continue;
    // Nog niet verstuurd
    if (betaalStatus === 'VERZONDEN') continue;
    // Check timing: 4 uur na verzending
    var tsIdx = headers.indexOf('laatsteUpdate');
    if (tsIdx !== -1) {
      var ts = rij[tsIdx];
      if (ts instanceof Date) {
        var uurOud = (new Date() - ts) / (1000 * 60 * 60);
        if (uurOud < BETAALMAIL_VERTRAGING) {
          Logger.log('Skip betaalmail: pas ' + uurOud.toFixed(1) + 'u oud, wacht op ' + BETAALMAIL_VERTRAGING + 'u — rij ' + (i+1));
          continue;
        }
      }
    }
    // Stuur betalingsmail — voordeel via JS berekening (niet sheet formule)
    try {
      var record = bouwRecordVanBerekeningRij_(rij, headers);
      var voordeelJaar1 = record.voordeel_jaar1 || 0;
      Logger.log('Rij ' + (i+1) + ' voordeel_jaar1 (JS): ' + voordeelJaar1);
      if (voordeelJaar1 < STRIPE_DREMPEL) {
        var idx = kolom_(headers, 'betalingsMailStatus');
        if (idx !== null) sheet.getRange(i + 1, idx + 1).setValue('NIET_RENDABEL');
        SpreadsheetApp.flush();
        continue;
      }
      verstuurBetalingsMail_(record);
      var betaalIdx = kolom_(headers, 'betalingsMailStatus');
      if (betaalIdx !== null) sheet.getRange(i + 1, betaalIdx + 1).setValue('VERZONDEN');
      SpreadsheetApp.flush();
      logRegel_('betaalmail_verzonden', rapportId, 'ok', record.mailAan);
      verzonden++;
    } catch(e) {
      var betaalIdx = kolom_(headers, 'betalingsMailStatus');
      if (betaalIdx !== null) sheet.getRange(i + 1, betaalIdx + 1).setValue('FOUT: ' + e.message);
      SpreadsheetApp.flush();
      Logger.log('Betaalmail fout rij ' + (i+1) + ': ' + e.message);
    }
  }
  Logger.log('Betalingsmails klaar \u2014 verzonden: ' + verzonden);
  return verzonden;
}
function verwerkAllesV2() {
  verwerkNieuweAanvragenV2();
}
function verwerkRapportenV2() {
  verwerkVrijgegevenDossiers();
}
function verwerkBetalingsMailsV2() {
  verwerkBetalingsmails();
}
// ============================================================================
// HULPFUNCTIES HTML
// ============================================================================
function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#039;');
}
function safeValue_(value) {
  if (value === null || value === undefined) return '';
  var text = String(value).trim();
  if (text === '#NAME?' || text === '#VALUE!' || text === '#REF!' ||
      text === '#N/A'   || text === '#ERROR!' || text === 'undefined' || text === 'null') return '';
  return text;
}
// ============================================================================
// PREVIEW EN TEST
// ============================================================================
function previewHtml(rijnummer) {
  if (!rijnummer || rijnummer < 3) rijnummer = 3;
  var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet   = ss.getSheetByName(TAB_BEREKENING);
  var data    = sheet.getDataRange().getValues();
  var headers = data[1].map(function(h) { return s_(h); });
  if (rijnummer > data.length) throw new Error('Rijnummer ' + rijnummer + ' bestaat niet');
  var rij    = data[rijnummer - 1];
  var record = bouwRecordVanBerekeningRij_(rij, headers);
  var html   = bouwRapportOutputHtml_(record);
  var blob   = Utilities.newBlob(html, 'text/html', 'preview_v2.8.32_rij' + rijnummer + '.html');
  var bestand = DriveApp.createFile(blob);
  Logger.log('Preview aangemaakt: ' + bestand.getUrl());
  return bestand.getUrl();
}
function testDossierNieuwste() {
  // Verwerk eerst nieuwe aanvragen naar Berekening_Rapport
  verwerkNieuweAanvragenV2();
  Utilities.sleep(2000);
  // Pak de laatste rij in Berekening_Rapport
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(TAB_BEREKENING);
  var lastRow = sheet.getLastRow();
  Logger.log('Laatste rij in Berekening_Rapport: ' + lastRow);
  return testDossier(lastRow);
}
function testDossier(rijnummer) {
  var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet   = ss.getSheetByName(TAB_BEREKENING);
  var data    = sheet.getDataRange().getValues();
  var headers = data[1].map(function(h) { return s_(h); });
  if (rijnummer < 2 || rijnummer > data.length) throw new Error('Rijnummer ' + rijnummer + ' bestaat niet.');
  var vrijIdx = kolom_(headers, 'rapportVrijgeven');
  if (vrijIdx !== null) { sheet.getRange(rijnummer, vrijIdx + 1).setValue('JA'); SpreadsheetApp.flush(); }
  var rij = sheet.getDataRange().getValues()[rijnummer - 1];
  return verwerkDossier_(rij, headers, rijnummer);
}
// ============================================================================
// TRIGGER EN SETUP
// ============================================================================
function installeerTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (fn === 'verwerkAllesV2' || fn === 'verwerkRapportenV2') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('verwerkAllesV2').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('verwerkRapportenV2').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('verwerkBetalingsMailsV2').timeBased().everyMinutes(15).create();
  Logger.log('Triggers geinstalleerd: verwerkAllesV2 + verwerkRapportenV2 + verwerkBetalingsMailsV2');
  return 'Triggers actief';
}
function setupV2() {
  var ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  var log = ['=== SENIOR ADVIES V2.8.32 SETUP ===', 'Gestart op: ' + new Date().toLocaleString('nl-NL'), ''];
  var vereisteTabs = ['Rapportaanvragen','Berekening_Rapport','Routekaarten','CAK_Parameters','CommunicatieLog'];
  var alleTabs = ss.getSheets().map(function(s) { return s.getName(); });
  vereisteTabs.forEach(function(naam) {
    log.push('  ' + (alleTabs.indexOf(naam) !== -1 ? '\u2713' : '\u2717 ONTBREEKT') + ' \u2014 ' + naam);
  });
  var triggerActief = ScriptApp.getProjectTriggers().some(function(t) { return t.getHandlerFunction() === 'verwerkAllesV2'; });
  if (!triggerActief) {
    ScriptApp.newTrigger('verwerkAllesV2').timeBased().everyMinutes(5).create();
    log.push('  \u2713 Trigger 1 (verwerkAllesV2) geinstalleerd');
  } else {
    log.push('  \u2713 Trigger 1 (verwerkAllesV2) was al actief');
  }
  var trigger2Actief = ScriptApp.getProjectTriggers().some(function(t) { return t.getHandlerFunction() === 'verwerkRapportenV2'; });
  if (!trigger2Actief) {
    ScriptApp.newTrigger('verwerkRapportenV2').timeBased().everyMinutes(5).create();
    log.push('  \u2713 Trigger 2 (verwerkRapportenV2) geinstalleerd');
  } else {
    log.push('  \u2713 Trigger 2 (verwerkRapportenV2) was al actief');
  }
  var trigger3Actief = ScriptApp.getProjectTriggers().some(function(t) { return t.getHandlerFunction() === 'verwerkBetalingsMailsV2'; });
  if (!trigger3Actief) {
    ScriptApp.newTrigger('verwerkBetalingsMailsV2').timeBased().everyMinutes(15).create();
    log.push('  \u2713 Trigger 3 (verwerkBetalingsMailsV2) geinstalleerd');
  } else {
    log.push('  \u2713 Trigger 3 (verwerkBetalingsMailsV2) was al actief');
  }
  var recoveryActief = ScriptApp.getProjectTriggers().some(function(t) { return t.getHandlerFunction() === 'checkRecoveryMails'; });
  if (!recoveryActief) {
    ScriptApp.newTrigger('checkRecoveryMails').timeBased().everyMinutes(15).create();
    log.push('  \u2713 Recovery trigger geinstalleerd');
  } else {
    log.push('  \u2713 Recovery trigger was al actief');
  }
  // Zorg dat betalingsMailStatus kolom bestaat
  var kolomResultaat = voegBetalingsKolomToe();
  log.push('  \u2713 Kolom betalingsMailStatus: ' + kolomResultaat);
  Logger.log(log.join('\n'));
  return log.join('\n');
}
// ============================================================================
// doPost
// ============================================================================
function doPost(e) {
  try {
    var params = {};
    if (e && e.postData && e.postData.contents) {
      var parts = e.postData.contents.split('&');
      parts.forEach(function(part) {
        var kv = part.split('=');
        if (kv.length === 2) {
          params[decodeURIComponent(kv[0])] = decodeURIComponent(kv[1].replace(/\+/g, ' '));
        }
      });
    }
    var action = params['action'] || '';
    if (action === 'rapport_aanvraag') {
      verwerkRapportAanvraag_(params);
    } else if (action === 'quickscan_lead') {
      verwerkQuickscanLead_(params);
    } else if (action === 'cta_click') {
      logCtaClick_(params);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    Logger.log('doPost fout: ' + err.message);
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}
// ============================================================================
// doGet
// ============================================================================
function doGet(e) {
  try {
    var params = e && e.parameter ? e.parameter : {};
    var action = params['action'] || '';
    if (action === 'cta_click') {
      logCtaClick_(params);
    } else if (action === 'quickscan_prefill') {
      return verwerkPrefillVerzoek_(params);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}
// ============================================================================
// Verwerk rapport aanvraag
// ============================================================================
function verwerkRapportAanvraag_(params) {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(TAB_RAPPORTAANVRAGEN);
  if (!sheet) throw new Error('Tab Rapportaanvragen niet gevonden');
  var headers   = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  var nieuweRij = new Array(headers.length).fill('');
  function zetKolom(naam, waarde) {
    var idx = headers.indexOf(naam);
    if (idx !== -1 && waarde) nieuweRij[idx] = waarde;
  }
  zetKolom('timestamp',           new Date());
  zetKolom('quickscanId',         params['quickscanId'] || '');
  zetKolom('formulierType',       params['formulierType'] || 'betaald_rapport_aanvraag');
  zetKolom('status',              'rapport_aanvraag');
  zetKolom('naam',                params['naam'] || '');
  zetKolom('geboorte',            params['geboorte'] || '');
  zetKolom('email',               params['email'] || '');
  zetKolom('telefoon',            params['telefoon'] || '');
  zetKolom('adres',               params['adres'] || '');
  zetKolom('postcode',            params['postcode'] || '');
  zetKolom('woonplaats',          params['woonplaats'] || '');
  zetKolom('aanhef',              params['aanhef'] || '');
  zetKolom('contactpersoonNaam',  params['contactpersoonNaam'] || '');
  zetKolom('contactpersoonEmail', params['contactpersoonEmail'] || '');
  zetKolom('partner',             params['partner'] || '');
  zetKolom('partnerThuis',        params['partnerThuis'] || '');
  zetKolom('partnerGeboorte',     params['partnerGeboorte'] || '');
  zetKolom('aowStatus',           params['aowStatus'] || '');
  zetKolom('opname',              params['opname'] || '');
  zetKolom('startdatumOpname',    params['startdatumOpname'] || '');
  zetKolom('zorgvorm',            params['zorgvorm'] || '');
  zetKolom('opnameDuur',          params['opnameDuur'] || '');
  zetKolom('bijdrageTypeBekend',  params['bijdrageTypeBekend'] || '');
  zetKolom('bijdrage',            params['bijdrage'] || '');
  zetKolom('inkomenPeiljaar',     params['inkomenPeiljaar'] || '');
  zetKolom('inkomenGedaald',      params['inkomenGedaald'] || '');
  zetKolom('redenInkomensdaling', params['redenInkomensdaling'] || '');
  zetKolom('inkomenActueel',      params['inkomenActueel'] || '');
  zetKolom('vermogen',            params['vermogen'] || '');
  zetKolom('vermogenActueel',     params['vermogenActueel'] || '');
  zetKolom('schuldenBox3',        params['schuldenBox3'] || '');
  zetKolom('woning',              params['woning'] || '');
  zetKolom('woningStatus',        params['woningStatus'] || '');
  zetKolom('cakBeschikking',      params['cakBeschikking'] || '');
  zetKolom('svbGeinformeerd',     params['svbGeinformeerd'] || '');
  zetKolom('bevestigd',           params['bevestigd'] || '');
  zetKolom('rapportVrijgeven',    'JA');
  sheet.appendRow(nieuweRij);
  SpreadsheetApp.flush();
  logRegel_('rapport_aanvraag_ontvangen', params['quickscanId'], 'ok', params['naam']);
}
// ============================================================================
// Verwerk quickscan lead
// ============================================================================
function verwerkQuickscanLead_(params) {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(TAB_RAPPORTAANVRAGEN);
  if (!sheet) throw new Error('Tab Rapportaanvragen niet gevonden');
  var headers   = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  var nieuweRij = new Array(headers.length).fill('');
  function zetKolom(naam, waarde) {
    var idx = headers.indexOf(naam);
    if (idx !== -1 && waarde) nieuweRij[idx] = waarde;
  }
  zetKolom('timestamp',        new Date());
  zetKolom('quickscanId',      params['quickscanId'] || '');
  zetKolom('formulierType',    'quickscan_lead');
  zetKolom('status',           'quickscan_lead');
  zetKolom('naam',             params['naam'] || '');
  zetKolom('email',            params['email'] || '');
  zetKolom('partner',          params['partner'] || '');
  zetKolom('partnerThuis',     params['partnerThuis'] || '');
  zetKolom('bijdrage',         params['eigenBijdrage'] || '');
  zetKolom('inkomenPeiljaar',  params['inkomenPeiljaar'] || '');
  zetKolom('vermogen',         params['vermogen'] || '');
  zetKolom('rapportVrijgeven', '');
  sheet.appendRow(nieuweRij);
  SpreadsheetApp.flush();
  logRegel_('quickscan_lead_ontvangen', params['quickscanId'], 'ok', params['naam']);
}
// ============================================================================
// Log CTA klik
// ============================================================================
function logCtaClick_(params) {
  var quickscanId = params['quickscanId'] || '';
  logRegel_('cta_click', quickscanId, 'ok', 'CTA geklikt');
  if (!quickscanId) return;
  try {
    var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet   = ss.getSheetByName(TAB_RAPPORTAANVRAGEN);
    var data    = sheet.getDataRange().getValues();
    var headers = data[1].map(function(h) { return s_(h); });
    var qIdx    = headers.indexOf('quickscanId');
    var ctaIdx  = headers.indexOf('ctaGeklikt');
    var recIdx  = headers.indexOf('recoveryMailStatus');
    for (var i = 2; i < data.length; i++) {
      if (s_(data[i][qIdx]) === quickscanId) {
        if (ctaIdx !== -1) sheet.getRange(i + 1, ctaIdx + 1).setValue('Ja');
        if (recIdx !== -1) sheet.getRange(i + 1, recIdx + 1).setValue('GEEN_MAIL_CTA_GEKLIKT');
        SpreadsheetApp.flush();
        break;
      }
    }
  } catch(err) { Logger.log('logCtaClick_ fout: ' + err.message); }
}
// ============================================================================
// Prefill verzoek
// ============================================================================
function verwerkPrefillVerzoek_(params) {
  var quickscanId = params['quickscanId'] || '';
  var callback    = params['callback'] || '';
  if (!quickscanId) {
    return ContentService.createTextOutput(callback + '(' + JSON.stringify({ ok: false }) + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet   = ss.getSheetByName(TAB_RAPPORTAANVRAGEN);
  var data    = sheet.getDataRange().getValues();
  var headers = data[1];
  var qIdx    = headers.indexOf('quickscanId');
  var gevonden = null;
  for (var i = 1; i < data.length; i++) {
    if (s_(data[i][qIdx]) === quickscanId) { gevonden = data[i]; break; }
  }
  if (!gevonden) {
    return ContentService.createTextOutput(callback + '(' + JSON.stringify({ ok: false }) + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  function val(naam) { var idx = headers.indexOf(naam); return idx !== -1 ? s_(gevonden[idx]) : ''; }
  var resultaat = { ok: true, data: {
    naam: val('naam'), email: val('email'), partner: val('partner'),
    partnerThuis: val('partnerThuis'), aowStatus: val('aowStatus'),
    bijdrageTypeBekend: val('bijdrageTypeBekend'), bijdrage: val('bijdrage'),
    inkomenPeiljaar: val('inkomenPeiljaar'), inkomen: val('inkomenPeiljaar'), vermogen: val('vermogen'),
  }};
  return ContentService.createTextOutput(callback + '(' + JSON.stringify(resultaat) + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
}
// ============================================================================
// RECOVERY MAIL
// FIX V2.8.29: geen besparingssuggestie voor alleenstaanden zonder voordeel.
//              Dubbel "mogelijk mogelijk" verwijderd.
// ============================================================================
function checkRecoveryMails() {
  var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet   = ss.getSheetByName(TAB_RAPPORTAANVRAGEN);
  if (!sheet) return;
  var data    = sheet.getDataRange().getValues();
  var headers = data[1].map(function(h) { return s_(h); });
  var nu            = new Date();
  var dertigMinuten = 30 * 60 * 1000;
  for (var i = 1; i < data.length; i++) {
    var rij = data[i];
    function val(naam) { var idx = headers.indexOf(naam); return idx !== -1 ? s_(rij[idx]) : ''; }
    if (val('formulierType') !== 'quickscan_lead') continue;
    if (val('recoveryMailStatus') === 'VERZONDEN') continue;
    if (val('recoveryMailStatus') === 'GEEN_MAIL_CTA_GEKLIKT') continue;
    if (val('ctaGeklikt') === 'Ja') continue;
    if (val('status') === 'rapport_aanvraag') continue;
    var qsId = val('quickscanId');
    if (qsId) {
      var sheetBer = ss.getSheetByName(TAB_BEREKENING);
      if (sheetBer) {
        var berData = sheetBer.getDataRange().getValues();
        var berHeaders = berData[1].map(function(h) { return s_(h); });
        var berQsIdx = berHeaders.indexOf('quickscanId');
        var berStatusIdx = berHeaders.indexOf('controleStatus');
        for (var b = 2; b < berData.length; b++) {
          if (s_(berData[b][berQsIdx]) === qsId && s_(berData[b][berStatusIdx]) === 'v2_verzonden') {
            Logger.log('Recovery skip: rapport al verzonden voor ' + qsId);
            updateRecoveryStatus_(sheet, i + 1, headers, 'GEEN_MAIL_RAPPORT_VERZONDEN');
            break;
          }
        }
      }
    }
    var huidigeStatus = sheet.getRange(i + 1, headers.indexOf('recoveryMailStatus') + 1).getValue();
    if (s_(huidigeStatus) === 'GEEN_MAIL_RAPPORT_VERZONDEN') continue;
    var timestamp = rij[headers.indexOf('timestamp')];
    if (!timestamp || !(timestamp instanceof Date)) continue;
    if (nu.getTime() - timestamp.getTime() < dertigMinuten) continue;
    try { verstuurRecoveryMail_(rij, headers, i + 1); }
    catch(err) {
      Logger.log('Recovery mail fout rij ' + (i+1) + ': ' + err.message);
      updateRecoveryStatus_(sheet, i + 1, headers, 'FAILED');
    }
  }
}
function verstuurRecoveryMail_(rij, headers, rijnummer) {
  function val(naam) { var idx = headers.indexOf(naam); return idx !== -1 ? s_(rij[idx]) : ''; }
  function n(naam) { return parseFloat(val(naam)) || 0; }
  var naam          = val('naam') || 'daar';
  var email         = val('email');
  var quickscanId   = val('quickscanId');
  var partner       = val('partner') === 'Ja';
  var partnerThuis  = val('partnerThuis') === 'Ja';
  var inkomen       = n('inkomenPeiljaar');
  var vermogen      = n('vermogen');
  var eigenBijdrage = n('bijdrage');
  if (!email) return;
  var toetsbedrag      = partnerThuis ? 73904 : 36952;
  var lageMin          = 212.60;
  var lageMax          = 1115.80;
  var vermBijt         = Math.max(0, (vermogen - toetsbedrag) * 0.04);
  var bijdrageplichtig = inkomen + vermBijt;
  var lageBijdrage     = Math.max(lageMin, Math.min(lageMax, bijdrageplichtig * 0.10 / 12));
  // FIX V2.8.29: alleen voordeel berekenen als partner thuis woont (BIJDRAGEVORM scenario)
  // Alleenstaanden hebben geen BIJDRAGEVORM voordeel via quickscan — geen bedrag tonen
  var scenarioB    = partner && partnerThuis && eigenBijdrage > lageBijdrage * 1.5;
  var voordeelMnd  = scenarioB ? Math.max(0, eigenBijdrage - lageBijdrage) : 0;
  var voordeelJaar = Math.round(voordeelMnd * 12);
  var rendabel     = scenarioB && voordeelJaar >= 150;
  // FIX V2.8.29: geen "mogelijk voordeel" tonen als er geen scenario is — neutrale tekst
  var heeftBedrag  = rendabel && voordeelJaar > 0;
  var bedragTekst  = heeftBedrag
    ? '\u20ac\u202f' + voordeelJaar.toLocaleString('nl-NL')
    : null;
  var ctaUrl = 'https://senior-advies.nl/uitgebreide-beoordeling-eigen-bijdrage-zorg-3c/'
    + (quickscanId ? '?quickscanId=' + encodeURIComponent(quickscanId) : '');
  // FIX V2.8.29: highlight blok alleen tonen als er een concreet bedrag is
  var highlightBlok = heeftBedrag
    ? '<div class="highlight">Onze eerste berekening laat zien dat u <strong>\u20ac\u202f' + voordeelJaar.toLocaleString('nl-NL') + '</strong> per jaar te veel betaalt.</div>'
    : '<div class="highlight">Op basis van uw gegevens zien wij een aanknopingspunt om uw eigen bijdrage te laten controleren.</div>';
  var htmlBody = '<!DOCTYPE html><html lang="nl"><head><meta charset="UTF-8"><style>'
    + 'body{margin:0;padding:0;background:#f4f8fa;font-family:Arial,Helvetica,sans-serif;color:#1f2933;}'
    + '.wrap{max-width:600px;margin:0 auto;background:#fff;}'
    + '.header{background:#004b64;padding:28px 32px;}'
    + '.header p{color:#fff;margin:0;font-size:13px;font-weight:800;}'
    + '.body{padding:32px;}h1{color:#004b64;font-size:22px;font-weight:800;margin:0 0 20px;}p{color:#253746;font-size:14px;line-height:1.65;margin:0 0 16px;}'
    + '.highlight{background:#f0f7fa;border-left:4px solid #ff532f;padding:14px 18px;margin:20px 0;font-size:15px;font-weight:700;color:#004b64;}'
    + '.cta-knop{display:inline-block;padding:16px 32px;background:#ff532f;color:#fff!important;text-decoration:none;border-radius:32px;font-size:16px;font-weight:800;}'
    + '.footer{padding:24px 32px;border-top:1px solid #dce9ee;font-size:12px;color:#6b7f8c;}'
    + '</style></head><body><div class="wrap">'
    + '<div class="header"><p>Senior Advies</p></div>'
    + '<div class="body">'
    + '<h1>Uw analyse staat klaar</h1>'
    + '<p>Beste ' + esc_(naam) + ',</p>'
    + '<p>U heeft onlangs een quickscan ingevuld op senior-advies.nl. Op basis van uw gegevens hebben wij een eerste beoordeling gemaakt van uw eigen bijdrage.</p>'
    + highlightBlok
    + '<p>Veel mensen laten dit liggen &#8212; simpelweg omdat ze niet weten dat ze recht hebben op een lagere bijdrage. Wij helpen u dat uit te zoeken.</p>'
    + '<p>Voor een volledig persoonlijk beoordelingsrapport hebben wij nog een aantal aanvullende gegevens nodig. Het invullen duurt minder dan 5 minuten.</p>'
    + '<p style="text-align:center;margin:28px 0;"><a href="' + ctaUrl + '" class="cta-knop">Vraag uw beoordeling aan &rarr;</a></p>'
    + '<p>Met vriendelijke groet,<br><strong>Arnout J. Punt</strong><br>Adviseur eigen bijdrage<br>Senior Advies</p>'
    + '</div><div class="footer">Senior Advies | Amstelveen | 020 463 2990 | intake@senior-advies.nl | senior-advies.nl</div>'
    + '</div></body></html>';
  GmailApp.sendEmail(email, 'Uw persoonlijke eerste analyse staat voor u klaar \u2014 Senior Advies', '', {
    htmlBody: htmlBody, name: 'Senior Advies', from: 'intake@senior-advies.nl', cc: INTERNAL_CC_EMAIL,
  });
  updateRecoveryStatus_(SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(TAB_RAPPORTAANVRAGEN), rijnummer, headers, 'VERZONDEN');
  logRegel_('recovery_mail_verzonden', quickscanId, 'ok', email);
}
function updateRecoveryStatus_(sheet, rijnummer, headers, status) {
  var idx = headers.indexOf('recoveryMailStatus');
  if (idx !== -1) { sheet.getRange(rijnummer, idx + 1).setValue(status); SpreadsheetApp.flush(); }
}
function installeerRecoveryTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'checkRecoveryMails') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkRecoveryMails').timeBased().everyMinutes(15).create();
  return 'Recovery trigger actief';
}
// ============================================================================
// HULPFUNCTIE: zet exacte CAK-formules op een specifieke rij
// ============================================================================
function zetFormulesToeOpRij_(sheet, r) {
  var formules = [
    '=CAK_Parameters!$B$4+(O'+r+'="Ja")*(CAK_Parameters!$C$4-CAK_Parameters!$B$4)',
    '=CAK_Parameters!$D$4',
    '=CAK_Parameters!$E$4+(O'+r+'="Ja")*(CAK_Parameters!$F$4-CAK_Parameters!$E$4)',
    '"PROGNOSE"',
    '=ALS((O'+r+'<>"Ja")*(AA'+r+'="Hoge eigen bijdrage");"HOOG";"LAAG")',
    '=12',
    '=ALS(L'+r+'>0;MAX(0;(L'+r+'-CT'+r+')*0.04);0)',
    '=MAX(CAK_Parameters!$K$4;ALS(H'+r+'>0;MIN(CAK_Parameters!$L$4;(H'+r+'+CZ'+r+')*0.1/12);0))',
    '=ALS(H'+r+'>0;MIN(CAK_Parameters!$M$4;MAX(0;(DC'+r+'-DD'+r+'+CZ'+r+')/12));0)',
    '=ALS(H'+r+'>0;ALS(H'+r+'<=38441;H'+r+'*0.8152;ALS(H'+r+'<=76817;38441*0.8152+(H'+r+'-38441)*0.6303;38441*0.8152+(76817-38441)*0.6303+(H'+r+'-76817)*0.505))-CU'+r+'-CV'+r+'-CAK_Parameters!$G$4;0)',
    '=ALS(DC'+r+'>0;MAX(0;(DC'+r+'-(CAK_Parameters!$I$4+(O'+r+'="Ja")*(CAK_Parameters!$J$4-CAK_Parameters!$I$4)))*0.25);0)',
    '=ALS(CX'+r+'="HOOG";DB'+r+';ALS(AB'+r+'>0;AB'+r+';DA'+r+'))',
    '=ALS(EN(AA'+r+'="Hoge eigen bijdrage";O'+r+'="Ja");"JA";"NEE")',
    '=ALS(DF'+r+'="JA";DA'+r+';DE'+r+')',
    '=ALS(DF'+r+'="JA";MAX(0;DE'+r+'-DG'+r+');0)',
    '=ALS(EN(ISGETAL(ZOEKEN("Gehuwden";Q'+r+'));O'+r+'="Ja");"JA";"NEE")',
    '=ALS(DI'+r+'="JA";(CAK_Parameters!$O$4-CAK_Parameters!$N$4)-MAX(0;(((H'+r+'+(CAK_Parameters!$O$4-CAK_Parameters!$N$4)*12)*0.8152-CU'+r+'-CAK_Parameters!$E$4-CAK_Parameters!$G$4-MAX(0;(((H'+r+'+(CAK_Parameters!$O$4-CAK_Parameters!$N$4)*12)*0.8152-CU'+r+'-CAK_Parameters!$E$4-CAK_Parameters!$G$4-CAK_Parameters!$I$4)*0.25))+CZ'+r+')/12)-DG'+r+');0)',
    '=ALS(DI'+r+'="JA";ALS(DJ'+r+'>0;"JA - omzetten voordelig";"NEE - niet omzetten");"N.v.t.")',
    '=ALS(O'+r+'="Ja";DA'+r+'*CY'+r+';(DA'+r+'*MIN(4;CY'+r+'))+(DE'+r+'*MAX(0;CY'+r+'-4)))',
    '=ALS(DF'+r+'="JA";DG'+r+'*CY'+r+';DL'+r+')',
    '=MAX(0;DL'+r+'-DM'+r+')',
    '=ALS(DN'+r+'>=150;"JA";"NEE")',
    '=ALS(EN(DH'+r+'=0;DJ'+r+'<=0);"Geen rendabele route gevonden.";ALS(EN(DH'+r+'>0;DJ'+r+'>0;DI'+r+'="JA");"Bijdragevorm en AOW. Jaar 1: "&TEKST(DN'+r+';"#.##0")&" euro.";ALS(DH'+r+'>0;"Bijdragevorm. Jaar 1: "&TEKST(DN'+r+';"#.##0")&" euro.";"AOW. Jaar 1: "&TEKST(DN'+r+';"#.##0")&" euro.")))',
  ];
  sheet.getRange(r, 98, 1, 23).setFormulas([formules]);
}

// ============================================================================
// TEST BETALINGSMAIL — eenmalig uitvoeren om te testen
// ============================================================================
function testBetalingsMail() {
  var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet   = ss.getSheetByName(TAB_BEREKENING);
  var data    = sheet.getDataRange().getValues();
  var headers = data[1].map(function(h) { return s_(h); });
  var verzonden = 0;
  for (var i = 2; i < data.length; i++) {
    var rij       = data[i];
    var rapportId = get_(rij, headers, 'rapportId');
    var status    = s_(rij[kolom_(headers, 'controleStatus') || 0]);
    if (status !== 'v2_verzonden') continue;
    var betaalStatus = get_(rij, headers, 'betalingsMailStatus');
    if (betaalStatus === 'VERZONDEN') continue;
    try {
      var record = bouwRecordVanBerekeningRij_(rij, headers);
      var voordeelJaar1 = record.voordeel_jaar1 || 0;
      Logger.log('Rij ' + (i+1) + ' — ' + rapportId + ' — voordeel (JS): ' + voordeelJaar1 + ' — betaalStatus: ' + betaalStatus);
      if (voordeelJaar1 < STRIPE_DREMPEL) {
        Logger.log('Skip: voordeel ' + voordeelJaar1 + ' < drempel ' + STRIPE_DREMPEL);
        continue;
      }
      verstuurBetalingsMail_(record);
      var betaalIdx = kolom_(headers, 'betalingsMailStatus');
      if (betaalIdx !== null) sheet.getRange(i + 1, betaalIdx + 1).setValue('VERZONDEN');
      SpreadsheetApp.flush();
      logRegel_('betaalmail_verzonden', rapportId, 'ok', record.mailAan);
      Logger.log('\u2713 Betalingsmail verzonden: ' + rapportId);
      verzonden++;
    } catch(e) {
      Logger.log('\u2717 Fout: ' + e.message);
    }
  }
  Logger.log('Klaar — ' + verzonden + ' betalingsmails verzonden');
  return verzonden;
}

// ============================================================================
// PRINTENBIND — AUTOMATISCHE DRUK-OPDRACHT
// ----------------------------------------------------------------------------
// Verstuurt na het opmaken van de PDF automatisch een druk-opdracht naar
// PrintenBind. Product: magazine (A4, vouwen & nieten). PDF wordt door
// PrintenBind opgehaald via een publieke URL (pdf_url uit PythonAnywhere).
//
// VEILIGHEID: zolang PRINTENBIND_ACTIEF = false wordt er NIETS besteld.
//   Zet pas op true wanneer:
//     1. je PrintenBind-account is goedgekeurd als VASTE KLANT, en
//     2. PythonAnywhere een publieke pdf_url teruggeeft.
//
// TOKEN: staat NIET in deze code. Zet hem eenmalig in de Script Properties:
//   Projectinstellingen -> Scripteigenschappen -> eigenschap toevoegen
//     naam:  PRINTENBIND_TOKEN
//     waarde: B48628E5-C7DD-425D-8413-F2E720041676
// ============================================================================

// --- INSTELLINGEN (pas hier aan, nergens anders) ---------------------------
var PRINTENBIND_ACTIEF      = false;   // <== HOOFDSCHAKELAAR. false = niets bestellen
var PRINTENBIND_BASIS_URL   = 'https://printenbind.nl/api/v1';
var PRINTENBIND_PRODUCT      = 'magazine';
var PRINTENBIND_SIZE         = 'a4';              // eindformaat A4 (van A3 gevouwen)
var PRINTENBIND_PRINTSIDE    = 'double';          // dubbelzijdig
var PRINTENBIND_COLOR        = 'all';             // alles in kleur
var PRINTENBIND_PAPIER       = '100';             // binnenwerk 100 grams
var PRINTENBIND_PAPIER_COVER = '250';             // omslag 250 grams
var PRINTENBIND_FINISHING    = 'folding-stapling';// vouwen & nieten
var PRINTENBIND_FOLDING      = 'middle_long';     // in het midden gevouwen
var PRINTENBIND_FINISH_EXTRA = 'none';            // geen ronde hoeken
var PRINTENBIND_COPIES       = 1;                 // aantal exemplaren
var PRINTENBIND_PRODUCTIE    = 'standard';        // 'fast' / 'standard' / 'budget'
var PRINTENBIND_LAND         = 'NL';

// --- TOKEN OPHALEN UIT SCRIPT PROPERTIES -----------------------------------
function printenbindToken_() {
  var token = PropertiesService.getScriptProperties().getProperty('PRINTENBIND_TOKEN');
  if (!token) throw new Error('PRINTENBIND_TOKEN ontbreekt in Scripteigenschappen.');
  return token;
}

// --- ADRES SPLITSEN IN STRAAT + HUISNUMMER ---------------------------------
// PrintenBind wil straat en huisnummer apart. In het dossier staat het adres
// als 1 veld (bv. "Hoofdstraat 12 A"). Deze functie splitst dat netjes.
function splitsAdres_(adresVolledig) {
  var adres = String(adresVolledig || '').trim();
  var m = adres.match(/^(.*?)\s+(\d+\s*[A-Za-z]?(?:\s*-\s*\d+)?)\s*$/);
  if (m) {
    return { straat: m[1].trim(), huisnummer: m[2].replace(/\s+/g, '').trim() };
  }
  // geen huisnummer gevonden: alles als straat, huisnummer leeg
  return { straat: adres, huisnummer: '' };
}

// --- EEN POST-CALL NAAR PRINTENBIND ----------------------------------------
function printenbindPost_(pad, payload) {
  var response = UrlFetchApp.fetch(PRINTENBIND_BASIS_URL + pad, {
    method:             'post',
    contentType:        'application/json',
    headers:            { 'Authorization': printenbindToken_() },
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true,
    followRedirects:    false,
    deadline:           60,
  });
  return {
    code:     response.getResponseCode(),
    body:     response.getContentText(),
    location: response.getHeaders()['location'] || response.getHeaders()['Location'] || '',
  };
}

// --- ORDER-ID UIT LOCATION-HEADER HALEN -------------------------------------
// PrintenBind geeft bv. "orders/113/articles/106195" terug. Hieruit pakken
// we 113 (order-id).
function orderIdUitLocation_(location) {
  var m = String(location || '').match(/orders\/(\d+)/);
  return m ? m[1] : '';
}

// ============================================================================
// HOOFDFUNCTIE: stuur druk-opdracht voor 1 dossier
// Geeft { ok, orderId, melding } terug. Gooit NOOIT door naar de aanroeper:
// een drukfout mag de rapportverwerking en de mail niet blokkeren.
// ============================================================================
function verstuurPrintEnBindOrder_(record, pdfUrl) {
  // Schakelaar uit? Dan netjes overslaan, geen fout.
  if (!PRINTENBIND_ACTIEF) {
    return { ok: false, orderId: '', melding: 'PRINTENBIND_ACTIEF staat uit \u2014 geen bestelling geplaatst' };
  }
  try {
    if (!pdfUrl) throw new Error('Geen publieke pdf_url beschikbaar voor PrintenBind');

    var rapportId  = record.rapportId || '';
    var naam       = record.naam || record.client_name || '';
    var adresSplit = splitsAdres_(record.adres);
    var postcode   = String(record.postcode || '').replace(/\s+/g, '').toUpperCase();
    var woonplaats = record.woonplaats || '';

    if (!adresSplit.straat || !postcode || !woonplaats) {
      throw new Error('Onvolledig afleveradres (straat/postcode/woonplaats) voor ' + rapportId);
    }

    // ---- STAP 1: artikel aanmaken (opent automatisch een nieuwe order) ----
    var artikel = {
      product:           PRINTENBIND_PRODUCT,
      add_file_method:   'url',
      file_url:          pdfUrl,
      file_overwrite:    true,
      production_method: PRINTENBIND_PRODUCTIE,
      order_reference:   rapportId,
      number:            1,
      copies:            PRINTENBIND_COPIES,
      size:              PRINTENBIND_SIZE,
      printside:         PRINTENBIND_PRINTSIDE,
      color:             PRINTENBIND_COLOR,
      papertype:         PRINTENBIND_PAPIER,
      papertype_cover:   PRINTENBIND_PAPIER_COVER,
      finishing:         PRINTENBIND_FINISHING,
      folding:           PRINTENBIND_FOLDING,
      finishing_extra:   PRINTENBIND_FINISH_EXTRA,
    };
    var r1 = printenbindPost_('/orders/articles', artikel);
    if (r1.code !== 201) {
      throw new Error('Artikel aanmaken mislukt (HTTP ' + r1.code + '): ' + r1.body.substring(0, 200));
    }
    var orderId = orderIdUitLocation_(r1.location);
    if (!orderId) throw new Error('Geen order-id ontvangen van PrintenBind');

    // ---- STAP 2: afleveradres instellen op de order ----
    var levering = {
      name_contact:      naam,
      street:            adresSplit.straat,
      streetnumber:      adresSplit.huisnummer,
      zipcode:           postcode,
      city:              woonplaats,
      country:           PRINTENBIND_LAND,
    };
    var r2 = printenbindPost_('/delivery/' + orderId, levering);
    if (r2.code !== 201 && r2.code !== 200 && r2.code !== 204) {
      throw new Error('Afleveradres instellen mislukt (HTTP ' + r2.code + '): ' + r2.body.substring(0, 200));
    }

    // ---- STAP 3: order afronden ----
    var r3 = printenbindPost_('/order/' + orderId + '/finish', {});
    if (r3.code !== 201 && r3.code !== 200 && r3.code !== 204) {
      throw new Error('Order afronden mislukt (HTTP ' + r3.code + '): ' + r3.body.substring(0, 200));
    }

    logRegel_('printenbind_besteld', rapportId, 'ok', 'order ' + orderId + ' \u2192 ' + naam);
    Logger.log('PrintenBind besteld: ' + rapportId + ' (order ' + orderId + ')');
    return { ok: true, orderId: orderId, melding: 'Order ' + orderId + ' geplaatst' };

  } catch (err) {
    var melding = err && err.message ? err.message : String(err);
    logRegel_('printenbind_fout', record.rapportId || '', 'fout', melding);
    Logger.log('PrintenBind fout: ' + (record.rapportId || '') + ' \u2014 ' + melding);
    // bewust GEEN throw: drukfout mag rapport + mail niet blokkeren
    return { ok: false, orderId: '', melding: melding };
  }
}// ============================================================================
// EENMALIG SCRIPT: voegt de twee PrintenBind-kolommen toe aan Berekening_Rapport
// Plak dit onderaan je Code.gs, sla op, kies bovenin de functie
// 'voegPrintenbindKolommenToe' en klik op Uitvoeren. Daarna mag dit blok weer
// weg (maar het mag ook blijven staan; het doet niets vanzelf).
// Veilig: bestaat een kolom al, dan wordt hij overgeslagen.
// ============================================================================
function voegPrintenbindKolommenToe() {
  var ss    = SpreadsheetApp.openById('1pUGmAWdS0YDtfJj09V8xZL2j2SUZA-Hmjpsc9QISrOE');
  var sheet = ss.getSheetByName('Berekening_Rapport');
  if (!sheet) { Logger.log('Tab Berekening_Rapport niet gevonden'); return 'FOUT: tab niet gevonden'; }

  // De kopregel staat op rij 2 (rij 1 is een titel/uitleg-rij)
  var kopRij  = 2;
  var laatsteKolom = sheet.getLastColumn();
  var koppen  = sheet.getRange(kopRij, 1, 1, laatsteKolom).getValues()[0];

  var teMaken = ['printenbindStatus', 'printenbindOrderId'];
  var toegevoegd = [];

  teMaken.forEach(function(naam) {
    // bestaat de kolom al? dan overslaan
    var bestaatAl = false;
    for (var i = 0; i < koppen.length; i++) {
      if (String(koppen[i]).trim() === naam) { bestaatAl = true; break; }
    }
    if (bestaatAl) {
      Logger.log('Kolom bestaat al, overgeslagen: ' + naam);
      return;
    }
    // voeg nieuwe kolom toe achteraan
    laatsteKolom = sheet.getLastColumn();
    var nieuweKolom = laatsteKolom + 1;
    sheet.getRange(kopRij, nieuweKolom).setValue(naam);
    toegevoegd.push(naam);
    Logger.log('Kolom toegevoegd: ' + naam + ' (kolom ' + nieuweKolom + ')');
  });

  SpreadsheetApp.flush();
  var melding = toegevoegd.length > 0
    ? 'Klaar — toegevoegd: ' + toegevoegd.join(', ')
    : 'Klaar — beide kolommen bestonden al, niets gewijzigd';
  Logger.log(melding);
  return melding;
}