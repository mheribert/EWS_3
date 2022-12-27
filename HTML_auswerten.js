var ver = 'V3.2010';
var fs = require('fs');

exports.rechne_wertungen = function (body, seite, runden_info) {
    return rechne_wertungen(body, seite, runden_info);
};

function rechne_wertungen(body, seite, runden_info) {
    var st_kl = runden_info[0].Startklasse;
    var rd = runden_info[0].Runde;
    var trunde = runden_info[0].RundeArt;
    var kl_punkte = Punkteverteilung(st_kl, trunde, rd);
    var Punkte = 0;
    var x;
    var verst;
    var back;
    switch (st_kl.substring(0, 3)) {
        case "BW_":
            if (typeof body["wgs" + seite] !== "undefined") {
                Punkte  = parseFloat(body["wgs" + seite]) * kl_punkte[0] / 10;  // Grundschritt
                Punkte += parseFloat(body["wtf" + seite]) * kl_punkte[2] / 10;  // Tanzfiguren
                Punkte += parseFloat(body["win" + seite]) * kl_punkte[4] / 10;  // Interpretation
                Punkte += parseFloat(body["wdp" + seite]) * kl_punkte[6] / 10;  // Dance Performance
                if (trunde === 'ER') {
                    Punkte += parseFloat(body["wbd" + seite]) * kl_punkte[1] / 10;  // Basic Dance
                    Punkte += parseFloat(body["wsi" + seite]) * kl_punkte[5] / 10;  // Spontane Interpretation
                } else {
                    Punkte += parseFloat(body["wgs" + seite]) * kl_punkte[1] / 10;
                    Punkte += parseFloat(body["win" + seite]) * kl_punkte[5] / 10;
                }
                if (runden_info[0].Runde.indexOf("r_schnell") > 0) {
                    Punkte = Punkte * 1.1;
                }
            } else {        // Auswertung für NewGuidelines
                if (typeof body["wng_tth" + seite] !== "undefined") {
                    var wr_kr = ["ng_ttd", "ng_tth", "ng_bda", "ng_dap", "ng_bdb", "ng_fta", "ng_fts", "ng_ftb", "ng_inf", "ng_ins", "ng_inb"];
                    kl_punkte = Punkteverteilung("BW_NG", "", "");
                    for (var k = 0; k < wr_kr.length; k++) {
                        Punkte += parseFloat(body["w" + wr_kr[k] + seite]) * kl_punkte[k];
                    }
                    if (runden_info[0].Runde.indexOf("r_schnell") > 0) {
                        Punkte = Punkte * 1.1;
                    }
                } else {
                    verst = new Array("wakrovw", "wjuniorvw", "wkleidungvw", "wtanzbereichvw", "wtanzzeitvw", "waufrufvw");
                    for (x = 0; x < verst.length; x++) {
                        if (body[verst[x] + seite] !== "") {
                            Punkte = Punkte + parseFloat(body[verst[x] + seite]);
                        }
                    }
                }
            }
            break;
        case "RR_":
            back = add_akro(body, seite);
            Punkte = back.pu;
            if (rd.indexOf("_Fu") > 0) { Punkte = 0; }

            if (typeof body['wsh' + seite] !== "undefined") {
                Punkte = parseFloat(body["wsh" + seite]) * kl_punkte[0] / 10;
                Punkte += parseFloat(body["wsd" + seite]) * kl_punkte[2] / 10;
                Punkte += parseFloat(body["wch" + seite]) * kl_punkte[4] / 10;
                if (trunde === 'ER') {       //    && runden_info[0].Runde != "Semi"
                    Punkte += parseFloat(body["wth" + seite]) * kl_punkte[1] / 10 || 0;
                    Punkte += parseFloat(body["wtd" + seite]) * kl_punkte[3] / 10 || 0;
                    Punkte += parseFloat(body["wtf" + seite]) * kl_punkte[5] / 10 || 0;
                    Punkte += parseFloat(body["wda" + seite]) * kl_punkte[6] / 10 || 0;
                } else {
                    Punkte += parseFloat(body["wsh" + seite]) * kl_punkte[1] / 10 || 0;
                    Punkte += parseFloat(body["wsd" + seite]) * kl_punkte[3] / 10 || 0;
                    Punkte += parseFloat(body["wch" + seite]) * kl_punkte[5] / 10 || 0;
                    Punkte += parseFloat(body["wch" + seite]) * kl_punkte[6] / 10 || 0;
                }
            }
            // Mehrkampf 
            if (typeof body['wmk_th' + seite] !== "undefined") {
                Punkte  = parseFloat(body["wmk_th" + seite]);
                Punkte += parseFloat(body["wmk_dh" + seite] || 0);
                Punkte += parseFloat(body["wmk_td" + seite] || 0);
                Punkte += parseFloat(body["wmk_dd" + seite] || 0);
            }
            break;
        case "F_B":
            if (typeof body['wtk' + seite] !== "undefined") {
                Punkte = parseFloat(body["wtk" + seite]) * kl_punkte[0] / 10;
                Punkte += parseFloat(body["wch" + seite]) * kl_punkte[1] / 10;
                Punkte += parseFloat(body["wtf" + seite]) * kl_punkte[2] / 10;
                Punkte += parseFloat(body["wab" + seite]) * kl_punkte[4] / 10;
                Punkte += parseFloat(body["waw" + seite]) * kl_punkte[5] / 10;
                Punkte += parseFloat(body["waf" + seite]) * kl_punkte[6] / 10;
            } else {
                verst = new Array("wsidebysidevw", "wakrovw", "whighlightvw", "wjuniorvw", "wkleidungvw", "wtanzbereichvw", "wtanzzeitvw", "waufrufvw");
                for (x = 0; x < verst.length; x++) {
                    if (body[verst[x] + seite] !== "") {
                        Punkte = Punkte + parseFloat(body[verst[x] + seite]);
                    }
                }
            }
//            Punkte = (Punkte * Form_abzuege(body["TP_ID" + seite], runden_info));
            break;
        case "F_R":
            back = add_akro(body, seite);
            Punkte = back.pu;
            if (typeof body['wtk' + seite] !== "undefined") {
                Punkte =  parseFloat(body["wtk" + seite]) * kl_punkte[0] / 10;
                Punkte += parseFloat(body["wch" + seite]) * kl_punkte[1] / 10;
                Punkte += parseFloat(body["wtf" + seite]) * kl_punkte[2] / 10;
                Punkte += parseFloat(body["wab" + seite]) * kl_punkte[4] / 10;
                Punkte += parseFloat(body["waw" + seite]) * kl_punkte[5] / 10;
                Punkte += parseFloat(body["waf" + seite]) * kl_punkte[6] / 10;
            }
            Punkte = Punkte * Form_abzuege(body["TP_ID" + seite], runden_info);
            break;
        case "BS_":
            switch (st_kl) {
                case "BS_BY_BJ":        // BRBV
                case "BS_BY_BE":
                case "BS_BY_BS":
                case "BS_BY_S1":
                    Punkte =  parseFloat(body["wgs" + seite]) * kl_punkte[0] / 10;  // Grundschritt
                    Punkte += parseFloat(body["wbd" + seite]) * kl_punkte[1] / 10;  // Basic Dancing, Lead & Follow, Harmonie
                    Punkte += parseFloat(body["wtf" + seite]) * kl_punkte[2] / 10;  // Tanzfiguren (einfache, highlight)
                    Punkte += parseFloat(body["win" + seite]) * kl_punkte[3] / 10;  // Interpretation (Figuren, spontane Interpretation)
                    break;
                case "BS_BW_BW":        // BWRRV
                case "BS_BW_SH":
                case "BS_F_BW_FO":
                case "BS_F_RR_EF":
                case "BS_F_RR_JF":
                case "BS_RR_BB":
                case "BS_RR_E1":
                case "BS_RR_J1":
                case "BS_RR_J2":
                case "BS_RR_S1":
                case "BS_RR_S2":
                    Punkte =  parseFloat(body["wth" + seite]);
                    Punkte += parseFloat(body["wtd" + seite]);
                    Punkte += parseFloat(body["wta" + seite]);   
                    Punkte += parseFloat(body["wak" + seite]);
                    Punkte -= parseFloat(body["wfe" + seite]);  
                    break;
                default:                //  DRBV
                    Punkte = parseFloat(body["wgs" + seite]);
                    break;
            }
            break;
        default:
            console.log("Fehler bei der Punkteberechnung Startklasse wurde nicht erkannt!!!!!!!!!!!!!!!!!");
            break;
    }
    if (Punkte < 0) {
        Punkte = 0;
    }
    return to_zahl(Punkte);
}

function add_akro(body, i) {
    var out = { pu:0, anz:0 };
    for (var z = 1; z <= 8; z++) {
        if (typeof body["wak" + i + z] !== "undefined") {
            out.pu += parseFloat(body["wak" + i + z]);
            out.anz++;
        }
    }
    return out;
}

exports.berechne_punkte = function (wertungen, runden_info, runde, wertungsrichter, f_name) {
    for (var s in runden_info) {
        if (runden_info[s].Rundennummer === runde) {
            PunkteOb = new Array();
            PunkteFt = new Array();
            PunkteAk = new Array();
            PunkteAe = new Object();
            runden_info[s].Punkte = 0;
            for (var i in wertungen) {
                for (var x in wertungen[i]) {
                    if (parseInt(x) === runden_info[s].TP_ID) {
                        if (wertungsrichter[i].WR_func === 'Ob') {
                            seite = wertungen[i][x].Seite;         //eins_zwei(runden_info[s].TP_ID, wertungen[i][x].cgi)
                            PunkteOb.push([i, x, 0, false, wertungen[i][x].cgi]);    //  WR, TP_ID, Punkte, Wertungen
                            switch (runden_info[0].Startklasse.substring(0, 3)) {
                                case "F_B":
                                case "BW_":
                                    PunkteOb[PunkteOb.length - 1][2] += to_zahl(wertungen[i][x].Punkte);
                                    break;
                                case "F_R":
                                case "RR_":
                                    if (runden_info[0].Startklasse ==="RR_S1" || runden_info[0].Startklasse === "RR_S2") {
                                        PunkteOb[PunkteOb.length - 1][2] += to_zahl(wertungen[i][x].Punkte) * -1
                                    } else {
                                        for (var aewr in wertungen[i][x]) {
                                            if ((aewr.substr(0, 4) === "tfl" + seite || aewr.substr(0, 6) === "tflak" + seite) && wertungen[i][x][aewr] != "") {
                                                PunkteAe[aewr] = wertungen[i][x][aewr];
                                                PunkteOb[PunkteOb.length - 1][2] += to_zahl(wertungen[i][x]["w" + aewr.substr(1, 8)]);
                                                if (wertungen[i][x][aewr].indexOf("P0") != -1) {
                                                    PunkteAe["w" + aewr.substr(3, 4)] = 0;
                                                    wertungen[i][x]["w" + aewr.substr(3, 4)] = 0;
                                                }
                                            }
                                        }
                                            
                                    }
                                break;
                                default:
                            }
                            write_back(wertungen, PunkteAe, x, i);  //  wertungen, PunkteAe, TP_ID, WR, filename
                        }
                    }
                }
            }
            for (i in wertungen) {
                PunkteAe = new Object();
                for (x in wertungen[i]) {
                    if (parseInt(x) === runden_info[s].TP_ID) {
                        wertungen[i][x].Punkte = rechne_wertungen(wertungen[i][x].cgi, wertungen[i][x].Seite, runden_info);
                        switch (wertungsrichter[i].WR_func) {
                            case 'Ft':
                            case 'X':
                            case 'MB':
                                PunkteFt.push([i, x, wertungen[i][x].Punkte, false, wertungen[i][x].cgi, wertungen[i][x].Seite, runden_info[0].Runde]);
                                break;
                            case 'MA':
                                PunkteAk.push([i, x, wertungen[i][x].Punkte, false, wertungen[i][x].cgi, wertungen[i][x].Seite, runden_info[0].Runde]);
                                break;
                            case 'Ak':
                                PunkteAk.push([i, x, wertungen[i][x].Punkte, false, wertungen[i][x].cgi, wertungen[i][x].Seite, runden_info[0].Runde]);
                                break;
                        }
                    }
                }
            }
            var st_kl = runden_info[0].Startklasse;
            runden_info[s].PunkteFt = get_mittel(PunkteFt, wertungen, st_kl);
            runden_info[s].PunkteAk = get_mittel(PunkteAk, wertungen, st_kl);
            runden_info[s].PunkteOb = get_mittel(PunkteOb, wertungen, st_kl);
            runden_info[s].Punkte = runden_info[s].PunkteFt + runden_info[s].PunkteAk - runden_info[s].PunkteOb;
            if (runden_info[s].Punkte < 0) { runden_info[s].Punkte = 0; }
            if (runden_info[s].Rundennummer === runde) {
                runden_info[s].berechnet = "Ok";
            }
        }
    }
    // in File schreiben
    var poststring;
    var wtext;
    for (w in wertungen) {
        for (pr in wertungen[w]) {
            if (parseInt(wertungen[w][pr].cgi.rh1) === runde) {
                poststring = '';
                wtext = '';
                for (i in wertungen[w][pr].cgi) {
                    poststring += i + '=' + wertungen[w][pr].cgi[i] + '&';
                }
                poststring += "Punkte" + wertungen[w][pr].Seite + '=' + wertungen[w][pr].Punkte + '&';
                if (typeof wertungen[w][pr].Punkte_err !== "undefined") {
                    poststring += 'Punkte_err' + wertungen[w][pr].Seite + '=' + wertungen[w][pr].Punkte_err + '&';
                } 
                if (typeof wertungen[w][pr].in === "undefined") {
                    poststring = poststring.substring(0, poststring.length - 1);
                } else {
                    poststring += 'wertung_in=' + wertungen[w][pr].in;
                }
                wtext += pr + ';' + w + ';' + poststring + '\r\n';
                wtext = wtext.replace(/TP_ID/g, "PR_ID");
                fs.appendFileSync(f_name, wtext, encoding = 'utf8');
            }
        }
        
    }
    // ObserverCheck in den Wertungen ändern
    function write_back(wertungen, PunkteAe, TP_ID, i) {
        var poststring;
        var p;
        for (wr in wertungen) {
            for (pr in wertungen[wr]) {
                if (wertungen[wr][pr].cgi.TP_ID1 === TP_ID || wertungen[wr][pr].cgi.TP_ID2 === TP_ID) {
                    for (p in PunkteAe) {
                        if (typeof wertungen[wr][pr].cgi[p] !== "undefined") {
                            if (p.substr(0, 3) !== "wak") {
                                wertungen[wr][pr].cgi[p] = wertungen[i][TP_ID][p];
                                wertungen[wr][pr].cgi["w" + p.substr(1, 10)] = to_zahl(wertungen[i][TP_ID]["w" + p.substr(1, 10)]);
                            } else {
                                wertungen[wr][pr].cgi["w" + p.substr(1, 10)] = to_zahl(PunkteAe[p]);
                            }

                        }
                    }
                }
            }
        }

    }
};

function get_mittel(avr, wertungen, st_kl) {
    var min;
    var max;
    var pu = 0;
    if (avr.length > 0) {
        avr.sort(function (a, b) {
            if (a[2] > b[2]) {
                return 1;
            }
            if (a[2] < b[2]) {
                return -1;
            }
            // a muss gleich b sein
            return 0;
        });
        switch (avr.length) {
            case 1:
            case 2:
                min = 1;
                max = avr.length;
                break;
            case 3:
                var rde = avr[0][6].substring(0, 4);
                min = 1;
                max = avr.length;
                if (rde.substring(0, 3) === "MK_") {
                    if (!(rde === 'MK_5' && (st_kl === "RR_J" || st_kl === "RR_S"))) {
                        max = 2;
                    } 
                } 
                break;
            case 4:
                min = 2;
                max = 3;
                break;
            case 5:
                min = 2;
                max = 4;
                if (st_kl.substring(0, 3) === "BW_") {          // bei Gleichheit der Punkte, kein aussortieren
                    if (avr[0][2] === avr[1][2]) { min--; }
                    if (avr[max][2] === avr[max - 1][2]) { max++; }
                }
                 break;
            case 6:
                min = 2;
                max = 5;
               break;
            case 7:
                min = 2;
                max = 6;
                if (st_kl.substring(0, 3) === "BW_") {          // bei Gleichheit der Punkte, kein aussortieren
                    if (avr[0][2] === avr[1][2]) { min--; }
                    if (avr[max][2] === avr[max - 1][2]) { max++; }
                }
                break;
            case 8:
                min = 3;
                max = 6;
                break;
            default:
                console.log("Fehler in der Anzahl der WR!");
        }
        var x;
        if (typeof avr[0][4]["wng_ttd" + avr[0][5]] === "undefined") {
            for (x = min - 1; x < max; x++) {
                pu = pu + parseFloat(avr[x][2]);
                wertungen[avr[x][0]][avr[x][1]].in = true;
            }
            return pu / (max - min + 1);
        } else {                     // ab hier Kategorien streichverfahren
            var kl_punkte = Punkteverteilung("BW_NG", "", "");
            var wr_kr = ["ng_ttd", "ng_tth", "ng_bda", "ng_dap", "ng_bdb", "ng_fta", "ng_fts", "ng_ftb", "ng_inf", "ng_ins", "ng_inb"];
            for (var kat = 0; kat < 11; kat++) {
                var kat_name = "w" + wr_kr[kat] + avr[0][5];
                var all = 0;
                for (x = min - 1; x < max; x++) {
                    all += parseFloat(avr[x][4][kat_name]);
                    wertungen[avr[x][0]][avr[x][1]].in = true;
                }
                var durchschnitt = all / (max - min + 1);
                var max_abw = 0;               // höchste differenz absolut
                var diff = [ 0, 0, 0, 0, 0, 0, 0];
                for (x = min - 1; x < max; x++) {
                    diff[x] = Math.abs(parseFloat(avr[x][4][kat_name]) - durchschnitt);
                    if (max_abw < Math.abs(parseFloat(avr[x][4][kat_name]) - durchschnitt)) {
                        max_abw = Math.abs(parseFloat(avr[x][4][kat_name]) - durchschnitt);
                    }
                } 
                var allrest = 0;
                var anzwrrest = 0;      // Reste addieren
                for (x = min - 1; x < max; x++) {
                    if ((diff[x] === max_abw) && (durchschnitt > 0 && max_abw !== 0)) {
                        wertungen[avr[x][0]][avr[x][1]][kat_name] = false;
                    } else {
                        allrest += parseFloat(avr[x][4][kat_name]);
                        anzwrrest++;
                    }
                }
                if ((durchschnitt > 0 && max_abw === 0) || anzwrrest === 0) {
                    pu += durchschnitt * kl_punkte[kat];
                } else {
                    if (anzwrrest !== 0) {              // daraus Mittelwert, keine division durch 0
                        pu += allrest / anzwrrest * kl_punkte[kat];
                    }
                }

            }
            for (x in wertungen) {
                wertungen[x][avr[0][1]]["Punkte_err"] = pu;
            }
            if (avr[0][6].indexOf("r_schnell") > 0) {
                pu = pu * 1.1;
            }
           return pu ;
        }
    } else {
        return parseFloat(0);
    }

}

function to_zahl(wert) {
    if (isNaN(wert) || wert === "") {
        return 0;
    } else {
        return parseFloat(wert);
    }
}

function Punkteverteilung(Startklasse, trunde, rd) {
    var punkte_verteilung;
    switch (Startklasse) {
        // Formationen
        case "F_RR_ST":         // Showteam
            punkte_verteilung = Array(15, 25, 20, 0, 7.5, 7.5, 25);
            break;
        case "F_RR_GF":         // Girl RR
            punkte_verteilung = Array(20, 20, 20, 0, 10, 10, 20);
            break;
        case "F_RR_LF":         // Lady RR
            punkte_verteilung = Array(20, 20, 20, 0, 10, 10, 20);
            break;
        case "F_RR_J":         // Jugend
            punkte_verteilung = Array(20, 20, 20, 0, 10, 10, 20);
            break;
        case "F_RR_Q":         //  Quattro RR
            punkte_verteilung = Array(20, 20, 20, 0, 10, 10, 20);
            break;
        case "F_RR_M":         //  Master RR
            punkte_verteilung = Array(20, 20, 20, 0, 10, 10, 20);
            break;
        case "F_BW_M":         //  Master BW
            punkte_verteilung = Array(25, 25, 25, 0, 9, 8, 8);
            break;
        // RR-Einzel
        case "RR_S":
            punkte_verteilung = Array(4.5, 4.5, 4.5, 4.5, 6.3, 6.3, 5.4);
            break;
        case "RR_J":
            punkte_verteilung = Array(6, 6, 6, 6, 8.4, 8.4, 7.2);
            break;
        case "RR_C":
            punkte_verteilung = Array(6, 6, 6, 6, 8.4, 8.4, 7.2);
            break;
        case "RR_B":
        case "RR_A":
            switch (trunde) {
                case "VR":
                case "ZR":
                    punkte_verteilung = Array(6.25, 6.25, 6.25, 6.25, 8.75, 8.75, 7.5);
                    break;
                case "ER":
                    punkte_verteilung = Array(4.375, 4.375, 4.375, 4.375, 6.125, 6.125, 5.25);
                    break;
            }                                                               //      35%    35%    30%
            if (rd === "Semi") { punkte_verteilung = Array(8.75, 8.75, 8.75, 8.75, 12.25, 12.25, 10.5); }
            break;
        // Boogie
        case "BW_MA":
            punkte_verteilung = Array(15, 7.5, 10, 0, 15, 10, 7.5);
            break;
        case "BW_MB":
            punkte_verteilung = Array(15, 7.5, 10, 0, 15, 10, 7.5);
            break;
        case "BW_JA":
            punkte_verteilung = Array(15, 7.5, 10, 0, 15, 10, 7.5);
            break;
        case "BW_SA":
            punkte_verteilung = Array(15, 7.5, 10, 0, 15, 10, 7.5);
            break;
        case "BW_SB":
            punkte_verteilung = Array(15, 7.5, 10, 0, 15, 10, 7.5);
            break;
        case "BW_NG":
//            punkte_verteilung = Array(1.5, 1.5, 1, 1, 1, 1, 1, 1, 3, 3, 3);
            punkte_verteilung = Array(1.5, 1.5, 1.5, 1.5, 1, 1, 1, 1, 2.5, 2.5, 2.5);
            break;
        // Breitensport Bayern Boogie
        case "BS_BY_BJ":
        case "BS_BY_BE":
        case "BS_BY_BS":
        case "BS_BY_S1":
            punkte_verteilung = Array(15, 15, 10, 25, 0, 0, 0);
            break;
        // Default
        default:
            punkte_verteilung = Array(10, 10, 10, 10, 10, 10, 10, 10, 10);
            break;
    }
    return punkte_verteilung;
}

function Form_abzuege(TP_ID, runden_info) {
    var st_kl = runden_info[0].Startklasse;
    for (var i in runden_info) {
        if (runden_info[i].TP_ID == TP_ID) {
            var f = Faktor_Formation_Abzuege(st_kl);
            return (100 - (f.max - runden_info[i].Anz_Taenzer) * f.faktor) / 100;
        }
    }
    return 0;
}

function Faktor_Formation_Abzuege(Startklasse) {
    var Faktor_Formation_Abzuege = new Object();
    switch (Startklasse) {
        case "F_RR_ST":         // Showteam
            Faktor_Formation_Abzuege.faktor = 0;
            Faktor_Formation_Abzuege.min = 4;
            Faktor_Formation_Abzuege.max = 16;
            break;
        case "F_RR_GF":         // Girl RR
            Faktor_Formation_Abzuege.faktor = 1.75;
            Faktor_Formation_Abzuege.min = 8;
            Faktor_Formation_Abzuege.max = 12;
            break;
        case "F_RR_LF":         // Lady RR
            Faktor_Formation_Abzuege.faktor = 1.25;
            Faktor_Formation_Abzuege.min = 8;
            Faktor_Formation_Abzuege.max = 16;
            break;
        case "F_RR_J":          // Jugend
            Faktor_Formation_Abzuege.faktor = 1.25;
            Faktor_Formation_Abzuege.min = 8;
            Faktor_Formation_Abzuege.max = 12;
            break;
        case "F_RR_Q":           // Quattro RR
            Faktor_Formation_Abzuege.faktor = 0;
            Faktor_Formation_Abzuege.min = 8;
            Faktor_Formation_Abzuege.max = 8;
            break;
        case "F_RR_M":           // Master RR
            Faktor_Formation_Abzuege.faktor = 1.25;
            Faktor_Formation_Abzuege.min = 8;
            Faktor_Formation_Abzuege.max = 12;
            break;
        case "F_BW_M":          // Master Boogie
            Faktor_Formation_Abzuege.faktor = 0;
            Faktor_Formation_Abzuege.min = 8;
            Faktor_Formation_Abzuege.max = 12;
            break;
        default:                // Falls Startklasse nicht gefunden
            Faktor_Formation_Abzuege.faktor = 0;
            Faktor_Formation_Abzuege.min = 0;
            Faktor_Formation_Abzuege.max = 0;
            break;
    }
    return Faktor_Formation_Abzuege;
}
