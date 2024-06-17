var ver = 'V3.2020' ;
var beamer_inhalt = new Object();
var HTML_Kopf = '';
var HTML_Inhalt = '';
var allranking;
var tp_id;

exports.inhalt = function (io) {
    io.emit('chat', { msg: 'beamer', bereich: 'beamer_kopf', cont: HTML_Kopf || '' });
    io.emit('chat', { msg: 'beamer', bereich: 'beamer_inhalt', cont: HTML_Inhalt || '' });
};

exports.beamer_seite = function (next) {
    var HTML_Seite;
    HTML_Seite = '<!DOCTYPE html>';
    if (typeof next === 'string') {     // next => beamer2  beamer3
        HTML_Seite += '<head><title>beamer' + next + '</title><meta http-equiv="expires" content="0">';
    } else {
        HTML_Seite += '<head><title>beamer</title><meta http-equiv="expires" content="0">';
    }
    HTML_Seite += '<link rel="stylesheet" href="EWS3.css">';
    HTML_Seite += '<script src="socket.io/socket.io.js"></script>';
    HTML_Seite += '<script src="beamod.js" ></script>';

    HTML_Seite += '</head><body style="height: 98%; font-family: Verdana;" id="beamer_seite">';

    HTML_Seite += '<table cellpadding="0" frame="void" class="tb1"><tr height = "20%" ><td><table width="100%">';
    HTML_Seite += '<tr><td id="beamer_bild" class="kopf" width="300px"><img src="logo.jpg" width="290" height="180" alt="DRBV"></td>';
    HTML_Seite += '<td class="kopf" width = "auto" id = "beamer_kopf">&nbsp;</td ></tr>';
    HTML_Seite += '</table></td></tr>';
    HTML_Seite += '<tr height="80%"><td><table style="width: 100%; float: left; " id="beamer_inhalt">';
    HTML_Seite += '<tr><td>&nbsp;</td></tr>';
    HTML_Seite += '</table></td></tr></table ></body ></html >';

    return HTML_Seite;

};

exports.beamer_runde = function (io, runden_info, runde, rd_ind) {
    // Kopf Text
    // var rd_ind = 0;
    if (typeof runden_info[0] === "undefined") { return; }
    if (runde <= runden_info[0].Tanzrunde_MAX) {
        /*for (i = 0; i < runden_info.length; i++) {
            if (runden_info[i].Rundennummer < runde) {
                rd_ind++;
            }
        }*/
        if (runde > runden_info[0].Tanzrunde_MAX) { return; }

        HTML_Kopf = runden_info[rd_ind].Turnier_Name + '<br>' + runden_info[rd_ind].Tanzrunde_Text;
        // Rundeninfo
        HTML_Inhalt = '<tr height="10%"><td colspan="2" class="runde">' + 'Runde ' + runden_info[rd_ind].Rundennummer + ' von ' + runden_info[rd_ind].Tanzrunde_MAX + '</td></tr>';
        // Startnummer(n)
        HTML_Inhalt += '<tr height="15%"><td class="stnr">' + runden_info[rd_ind].Startnr + '</td>';
        if (runden_info[rd_ind].PpR === 2) {
            HTML_Inhalt += '<td class="stnr">' + runden_info[rd_ind + 1].Startnr + '</td>';
        }
        HTML_Inhalt += '</tr>';
        // Paar(e), Team, Verein
        HTML_Inhalt += '<tr height="65%">';
        if (runden_info[rd_ind].Name_Team === null) {        // Einzel
            HTML_Inhalt += '<td class="tzer">' + runden_info[rd_ind].Dame + '<br>' + runden_info[rd_ind].Herr + '<br><p class="tver">' + runden_info[rd_ind].Verein_Name + '</p></td>';
            if (runden_info[rd_ind].PpR === 2) {
                HTML_Inhalt += '<td class="tzer">' + runden_info[rd_ind + 1].Dame + '<br>' + runden_info[rd_ind + 1].Herr + '<br><p class="tver">' + runden_info[rd_ind + 1].Verein_Name + '</p></td>';
            }
        } else {                                    // Formationen
            HTML_Inhalt += '<td class="tzer">' + runden_info[rd_ind].Name_Team + '<br><p class="tver">' + runden_info[rd_ind].Verein_Name + '</p></td>';
        }
        HTML_Inhalt += '</tr>';
        //WR-Info
        HTML_Inhalt += '<tr height="10%"><td colspan="2" align="center"><div class="wr_status" id="beamer_wrinfo">&nbsp;</div></td></tr>';
        io.emit('chat', { msg: 'beamer', bereich: 'beamer_kopf', cont: HTML_Kopf });
        io.emit('chat', { msg: 'beamer', bereich: 'beamer_inhalt', cont: HTML_Inhalt });
    }
};

exports.beamer_zeitplan = function (io, connection, ab_rtid) {
    connection
        .query('SELECT RT.RT_ID, RT.Turniernr, RT.Rundenreihenfolge, Startklasse_text, Rundentext,  Format([Startzeit],"Short Time") AS Zeit FROM Tanz_Runden INNER JOIN (Rundentab AS RT LEFT JOIN Startklasse ON RT.Startklasse = Startklasse.Startklasse) ON Tanz_Runden.Runde = RT.Runde WHERE RT.Rundenreihenfolge < 999 ORDER BY RT.Rundenreihenfolge;')
        .on('done', function (data) {
            // Kopf Text
            var beginn = false;
            HTML_Kopf = 'Zeitplan';
            // Rundeninfo
            HTML_Inhalt = '<tr height="100%"><td><table style="width: 100%; float: left; ">';
            HTML_Inhalt += '<thead><tr class="runden" role="row"><th style="width: 200px; padding-left:60px; " colspan="1" rowspan="1" class="sorting">Beginn</th><th style="width: auto;" colspan="1" rowspan="1" class="sorting">Runde</th></tr></thead>';
            HTML_Inhalt += '<tbody style="font-size: 2.5vw;">';
            for (var i in data) {
                if (typeof ab_rtid === undefined || ab_rtid === "") {
                    beginn = true;
                } else {
                    if (data[i].RT_ID.toString() === ab_rtid) {
                        beginn = true;
                    }
                }
                if (beginn === true && data[i].Rundentext !== "Letzte Startkartenabgabe" && data[i].Rundentext !== "Wertungsrichterbesprechung") {
                    HTML_Inhalt += '<tr class="odd" ><td style="padding-left:60px;" >' + data[i].Zeit + '</td><td>' + data[i].Rundentext + ' ' + (data[i].Startklasse_text || "") + '</td></tr>';
                }
            }
            io.emit('chat', { msg: 'beamer', bereich: 'beamer_kopf', cont: HTML_Kopf });
            io.emit('chat', { msg: 'beamer', bereich: 'beamer_inhalt', cont: HTML_Inhalt });
        });
};

exports.beamer_ranking = function (io, runden_info, runde) {
    if (!runden_info[0].ranking_anzeige ) { return; }
    var ratings = new Object();
    var temp = new Object;
    var ko_rde = new Object();
    var rde; 
    var anz = 0;
    // Kopf Text
    HTML_Kopf = runden_info[0].Turnier_Name  + '<br>' + runden_info[0].Tanzrunde_Text;
	// nur getanzte Runden
    var sum1;
    var sum2;
    for (var p in runden_info) {
        runden_info[p].Platz = "";
        rde = runden_info[p].Rundennummer || 0;
        if (runden_info[p].Rundennummer <= runde && runden_info[p].nochmal === false) {
            ratings[anz] = runden_info[p];
            ratings[anz].rd_info = p;
            sum1 = parseFloat(ratings[anz].Punkte);
            if (ratings[anz].ersteRunde !== null) {
                sum1 += parseFloat(ratings[anz].ersteRunde);
            }
            ratings[anz].summe = sum1;
             // Rundensieger für KO-Runde
            if (ko_rde[rde] < sum1 || ko_rde[rde] === undefined) {
                ko_rde[rde] = sum1;
                ko_rde["p"+ rde] = p;
            }
            anz++;
        }
    }
    // KO-Runde offset 
    if (runden_info[0].Runde.indexOf("KO_r") > -1) {
        for (p in ko_rde) {
            if (p.substring(0, 1) === "p") {
                runden_info[ko_rde[p]].kosieger = true;
            }
        }
    }   

    // sortieren nach punkten
    anz--;
    for (var s = 0; s < anz; s++) {
        for (var t = 0; t < anz; t++) {
            sum1 = parseFloat(ratings[t].summe) + (ratings[t].kosieger === true ? 2000 : 0);
            sum2 = parseFloat(ratings[t + 1].summe + (ratings[t + 1].kosieger === true ? 2000 : 0));
            if (sum1 < sum2) {
                temp = ratings[t];
                ratings[t] = ratings[t + 1];
                ratings[t + 1] = temp;
            }
        }
    } 
    // bei ersten Durchlauf erfolgt keine Schleife
    var pu = 0;
    var HTML_class;
    var platz = 0;
    var punkte = 0;
    var linie = false;
    var max_paare = 8;
    HTML_Inhalt = make_thead('') + '<tbody>';
    allranking = new Object();

    for (p in ratings) {
        if ((ratings[p].ersteRunde + ratings[p].Punkte) !== punkte) {
            platz = parseInt(p) + 1;
        //    if (p === ratings[p].NextPaare && ratings[p].NextPaare !== null) {
        //        HTML_Inhalt += '<tr class="trenn"><td>&nbsp;</td> <td>&nbsp;</td> <td>&nbsp;</td> <td>&nbsp;</td> </tr>';
        //        HTML_class = '<tr class="raus">';
        //    }
        }
        allranking[parseInt(p) + 1] = ratings[p].rd_info;
        runden_info[ratings[p].rd_info].Platz = platz;
        if (platz <= max_paare || ratings[p].Rundennummer === runde) {
            if (platz > max_paare && !linie) {
                linie = true;
                HTML_Inhalt += '<tr><td colspan="7">&nbsp;</td></tr>';
            }
            HTML_class = '<tr class="weiter">';
            if (ratings[p].Rundennummer === runde) {
                HTML_class = HTML_class.replace('>', ' style = "background-color:#cfc;">');
                HTML_Inhalt += HTML_class + '<td>*' + platz + '&nbsp;</td>';
            } else {
                HTML_Inhalt += HTML_class + '<td>' + platz + '&nbsp;</td>';
            }
            HTML_Inhalt += '<td>' + ratings[p].Startnr + '</td>';
            if (ratings[p].Name_Team === null) {
                HTML_Inhalt += '<td class="text_left">' + ratings[p].Dame + ' - ' + ratings[p].Herr + '</td>';
            } else {
                HTML_Inhalt += '<td class="text_left">' + ratings[p].Name_Team + '</td>';
            }
 //           if (ratings[p].punkte_anzeige === true) {       // keine Punkte anzeigen
                if (ratings[p].ersteRunde !== null) {
                    HTML_Inhalt += '<td style="font-size:1.6vw;">' + ratings[p].ersteRunde.toFixed(2) + ' + ' + ratings[p].Punkte.toFixed(2) + '</td>';
                }
                punkte = ratings[p].ersteRunde + ratings[p].Punkte;
                HTML_Inhalt += '<td>' + (ratings[p].ersteRunde + ratings[p].Punkte).toFixed(2).replace('.', ',') + ins_strafe(ratings[p]) + '</td>';
//            }
            HTML_Inhalt += '</tr>'
        }
    }
    HTML_Inhalt += '</tbody>';
    allranking.lenght = s + 1;
    io.emit('chat', { msg: 'beamer', bereich: 'beamer_kopf', cont: HTML_Kopf });
    io.emit('chat', { msg: 'beamer', bereich: 'beamer_inhalt', cont: HTML_Inhalt });

    function ins_strafe(strafe) {
        var back = "";
        if (strafe.a20 === true) {
            back = "&nbsp;<b>A</b>";
        }
        if (strafe.z20 === true) {
            back += "&nbsp;<b>Z</b>";
        }
        return back;
    }

};

exports.beamer_siegerehrung = function (io, connection, rt_id, Platz) {
    connection								
        .query('SELECT * FROM View_Rundenablauf WHERE RT_ID =' + rt_id + ' ORDER BY Platz, Startnr;')
        .on('done', function (data) {
            var HTML_Kopf = 'Siegerehrung<br>' + data[0].Tanzrunde_Text;
            var HTML_Inhalt = make_thead(data[0].Runde) + '<tbody>';
            var cl = '';
            for (var p in data) { 
                if (Platz &&  parseInt(Platz) > data[p].Platz) { // für getaktetes Anzeigen
                     cl = 'style="visibility :hidden;"';
                } else {
                    cl = '';
                }
                HTML_Inhalt += '<tr class="weiter"  ' + cl + '><td>' + data[p].Platz + '&nbsp;</td>';
                if (data[p].Name_Team === null) {
                    HTML_Inhalt += '<td>' + data[p].Startnr + '</td><td class="text_left">' + data[p].Dame + ' - ' + data[p].Herr + '</td>';
                } else {
                    HTML_Inhalt += '<td>' + data[p].Startnr + '</td><td class="text_left">' + data[p].Name_Team + '</td>';
                }
                data[p].jetztRunde = data[p].jetztRunde || 0;
                punkte = data[p].jetztRunde.toFixed(2);
                HTML_Inhalt += '<td class="pkte">' + punkte.replace('.', ',') + '</td></tr>';
            }
            HTML_Inhalt += '</tbody>';

            io.emit('chat', { msg: 'beamer', bereich: 'beamer_kopf', cont: HTML_Kopf });
            io.emit('chat', { msg: 'beamer', bereich: 'beamer_inhalt', cont: HTML_Inhalt });
        });
};

exports.beamer_allranking = function (io, runde, runden_info) {
    if (allranking === undefined) { return; }

    if (runde + 1> allranking.lenght) {
        io.emit('chat', { msg: 'beamer', bereich: 'beamer_inhalt', cont: HTML_Inhalt });
        return;
    }
    var HTML_Seite = make_thead('') + '<tbody>';

    for (p = runde; p < runde + 8; p++) {
        if (p + 1 > allranking.lenght) {
            break;
        }
        HTML_Seite += '<tr class="weiter"><td>' + runden_info[allranking[p + 1]].Platz + '&nbsp;</td>';
        HTML_Seite += '<td>' + runden_info[allranking[p + 1]].Startnr + '</td>';
        if (runden_info[allranking[p + 1]].Name_Team === null) {
            HTML_Seite += '<td class="text_left">' + runden_info[allranking[p + 1]].Dame + ' - ' + runden_info[allranking[p + 1]].Herr + '</td>';
        } else {
            HTML_Seite += '<td class="text_left">' + runden_info[allranking[p + 1]].Name_Team + '</td>';
        }
        if (runden_info[allranking[p + 1]].ersteRunde !== null) {
            HTML_Seite += '<td style="font-size:1.6vw;">' + runden_info[allranking[p + 1]].ersteRunde.toFixed(2) + ' + ' + runden_info[allranking[p + 1]].Punkte.toFixed(2) + '</td>';
        }
        punkte = (runden_info[allranking[p + 1]].summe).toFixed(2);
        HTML_Seite += '<td>' + punkte.replace('.', ',') + '</td></tr>';
    }
    HTML_Seite += '</tbody>';

    io.emit('chat', { msg: 'beamer', bereich: 'beamer_inhalt', cont: HTML_Seite });
        setTimeout(function () {
            exports.beamer_allranking(io, p, runden_info);
        }, 4000);

};

exports.beamer_stellprobe = function (io, connection, teams, title) {
    tp_id = teams.split(';');
    connection
        .query('SELECT * FROM paare WHERE TP_ID =' + tp_id[0] + ';')
        .on('done', function (data) {
            HTML_Kopf = 'Stellprobe<br>' + title ;
            HTML_Inhalt = '<tr class="runde"><td height = "25hv" width="20%">Jetzt:</td><td id = "jetzt">';
            if (tp_id[0] === '-1') {
                HTML_Inhalt += '<strong>Pause</strong></td></tr>';
            } else {
                HTML_Inhalt += '<strong>' + data[0].Name_Team + '</strong><br>' + data[0].Verein_Name + '</td></tr >';
            }
            connection
                .query('SELECT * FROM paare WHERE TP_ID =' + tp_id[1] + ';')
                .on('done', function (data) {

                    HTML_Inhalt += '<tr"><td height = "30hv" colspan="2" id="beamer_minute" style="text-align:center; font-size:7vw;">&nbsp;</td></tr>';
                    HTML_Inhalt += '<tr class="runde"><td height = "25hv" >Danach:</td><td>';
                    if (tp_id[1] === '-1') {
                        HTML_Inhalt += 'Pause</td></tr>';
                    } else {
                        HTML_Inhalt += '<strong>' + data[0].Name_Team + '</strong><br>' + data[0].Verein_Name + '</td></tr>';
                    }

                    io.emit('chat', { msg: 'beamer', bereich: 'beamer_kopf', cont: HTML_Kopf });
                    io.emit('chat', { msg: 'beamer', bereich: 'beamer_inhalt', cont: HTML_Inhalt });
                });
        });
};

function make_thead(rde) {
    var t_head = '<thead><tr><th style="width: 90px; font-size: 35px;" class="sorting text_center">Platz</th>';
    t_head += '<th style="width: 80px; font-size: 35px;" class="sorting text_center">&nbsp;StNr.&nbsp;</th>';
    t_head += '<th style="width: auto; font-size: 35px;" class="sorting">Paar</th>';
    if (rde ==="MK_5_TNZ") {
        t_head += '<th style="width: auto; font-size: 35px; text-align: center;" class="sorting">Summe aller<br>Pl&auml;tze</th></tr></thead>';
    } else {
        t_head += '<th style="width: auto; font-size: 35px; text-align: center;" class="sorting">Punkte</th></tr></thead>';
    }
    return t_head;
}
