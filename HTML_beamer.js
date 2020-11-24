var ver           = 'V3.1.16';
var beamer_inhalt = new Object();

exports.inhalt = function () {
    return beamer_inhalt;
};

exports.beamer_seite = function () {
    var HTML_Seite;

    HTML_Seite = '<!DOCTYPE html>';
    HTML_Seite += '<head><title>beamer</title><meta http-equiv="expires" content="0">';

    HTML_Seite += '<link rel="stylesheet" href="EWS3.css">';

    HTML_Seite += '<script src="socket.io/socket.io.js"></script>';
    HTML_Seite += '<script src="EWS3.js" ></script>';

    HTML_Seite += '</head><body style="height: 98%; font-family: Verdana;" id="beamer_seite">';

    HTML_Seite += '<table cellpadding="0" frame="void" class="tb1"><tr height = "20%" ><td><table width="100%">';
    HTML_Seite += '<tr><td class="kopf" width="300px"><img src="logo.jpg" width="290" height="180" alt="DRBV"></td>';
    HTML_Seite += '<td class="kopf" width = "auto" id = "beamer_kopf">&nbsp;</td ></tr>';
    HTML_Seite += '</table></td></tr>';
    HTML_Seite += '<tr height="80%"><td><table style="width: 100%; float: left; " id="beamer_inhalt">';
    HTML_Seite += '<tr><td>&nbsp;</td></tr>';
    HTML_Seite += '</table></td></tr></table ></body ></html >';

    return HTML_Seite;

};

exports.beamer_runde = function (io, runden_info, runde) {
    // Kopf Text
    var rd_ind = 0;
    if (runde <= runden_info[0].Tanzrunde_MAX) {
        for (i = 1; i < runde; i++) {
            rd_ind += parseInt(runden_info[i].PpR);
        }
        if (runde > runden_info[0].Tanzrunde_MAX) { return; }

        var HTML_Kopf = runden_info[rd_ind].Turnier_Name + '<br>' + runden_info[rd_ind].Tanzrunde_Text;
        // Rundeninfo
        var HTML_Inhalt = '<tr height="10%"><td colspan="2" class="runde">' + 'Runde ' + runden_info[rd_ind].Rundennummer + ' von ' + runden_info[rd_ind].Tanzrunde_MAX + '</td></tr>';
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
        beamer_inhalt = { msg: 'beamer', kopf: HTML_Kopf, inhalt: HTML_Inhalt };
        io.emit('chat', { msg: 'beamer', kopf: HTML_Kopf, inhalt: HTML_Inhalt });
    }
};

exports.beamer_zeitplan = function (io, connection, ab_rtid) {
    connection
        .query('SELECT RT.RT_ID, RT.Turniernr, RT.Rundenreihenfolge, Startklasse_text, Rundentext,  Format([Startzeit],"Short Time") AS Zeit FROM Tanz_Runden INNER JOIN (Rundentab AS RT LEFT JOIN Startklasse ON RT.Startklasse = Startklasse.Startklasse) ON Tanz_Runden.Runde = RT.Runde WHERE RT.Rundenreihenfolge < 999 ORDER BY RT.Rundenreihenfolge;')
        .on('done', function (data) {
            // Kopf Text
            var beginn = false;
            var HTML_Kopf = 'Zeitplan';
            // Rundeninfo
            var HTML_Inhalt = '<tr height="100%"><td><table style="width: 100%; float: left; ">';
            HTML_Inhalt += '<thead><tr class="runden" role="row"><th style="width: 250px; padding-left:80px; " colspan="1" rowspan="1" class="sorting">Beginn</th><th style="width: auto;" colspan="1" rowspan="1" class="sorting">Runde</th></tr></thead>';
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
                    HTML_Inhalt += '<tr class="odd" ><td style="padding-left:100px;" >' + data[i].Zeit + '</td><td>' + data[i].Rundentext + ' ' + (data[i].Startklasse_text || "") + '</td></tr>';
                }
            }
            beamer_inhalt = { msg: 'beamer', kopf: HTML_Kopf, inhalt: HTML_Inhalt };
            io.emit('chat', { msg: 'beamer', kopf: HTML_Kopf, inhalt: HTML_Inhalt });
        });
};

exports.beamer_ranking = function (io, runden_info, runde) {
    if (!runden_info[0].ranking_anzeige ) { return; }
    var ratings = new Object();
    var temp = new Object;
    var anz = 0;
    // Kopf Text
    var HTML_Kopf = runden_info[0].Turnier_Name  + '<br>' + runden_info[0].Tanzrunde_Text;
	// nur getanzte Runden
    var sum1;
    var sum2;
    for (var p in runden_info) {
        runden_info[p].Platz = "";
        if (runden_info[p].Rundennummer <= runde && runden_info[p].nochmal === false) {
            ratings[anz] = runden_info[p];
            ratings[anz].rd_info = anz; 
            sum1 = parseFloat(ratings[anz].Punkte);
            if (ratings[anz].ersteRunde !== null) {
                sum1 += parseFloat(ratings[anz].ersteRunde);
            }
            ratings[anz].summe = sum1;
            anz++;
        }
    }
    // sortieren nach punkten
    anz--;
    for (var s = 0; s < anz; s++) {
        for (var t = 0; t < anz; t++) {
            sum1 = parseFloat(ratings[t].summe);
            sum2 = parseFloat(ratings[t + 1].summe);
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
    var HTML_Inhalt = make_thead() + '<tbody>';

    for (p in ratings) {
        if (fix2(ratings[p].ersteRunde + ratings[p].Punkte) !== punkte) {
            platz = parseInt(p) + 1;
        //    if (p === ratings[p].NextPaare && ratings[p].NextPaare !== null) {
        //        HTML_Inhalt += '<tr class="trenn"><td>&nbsp;</td> <td>&nbsp;</td> <td>&nbsp;</td> <td>&nbsp;</td> </tr>';
        //        HTML_class = '<tr class="raus">';
        //    }
        }
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
            if (ratings[p].ersteRunde !== null) {
                HTML_Inhalt += '<td style="font-size:1.6vw;">' + fix2(ratings[p].ersteRunde) + ' + ' + fix2(ratings[p].Punkte) + '</td>';
            }
            punkte = fix2(ratings[p].ersteRunde + ratings[p].Punkte);
            HTML_Inhalt += '<td>' + punkte + '</td></tr>';
        }
    }
    HTML_Inhalt += '</tbody>';

    beamer_inhalt = { msg: 'beamer', kopf: HTML_Kopf, inhalt: HTML_Inhalt };
    io.emit('chat', { msg: 'beamer', kopf: HTML_Kopf, inhalt: HTML_Inhalt });
};

exports.beamer_siegerehrung = function (io, connection, rt_id, Platz) {
    connection								
        .query('SELECT * FROM View_Rundenablauf WHERE RT_ID =' + rt_id + ' ORDER BY Platz, Startnr;')
        .on('done', function (data) {
            var HTML_Kopf = 'Siegerehrung<br>' + data[0].Tanzrunde_Text;
            var HTML_Inhalt = make_thead() + '<tbody>';
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
                punkte = fix2(data[p].jetztRunde);
                HTML_Inhalt += '<td>' + punkte + '</td></tr>';
            }
            HTML_Inhalt += '</tbody>';

            beamer_inhalt = { msg: 'beamer', kopf: HTML_Kopf, inhalt: HTML_Inhalt };
            io.emit('chat', { msg: 'beamer', kopf: HTML_Kopf, inhalt: HTML_Inhalt });
        });
};

function make_thead() {
    var t_head = '<thead><tr><th style="width: 90px; font-size: 35px;" class="sorting text_center">Platz</th>';
    t_head += '<th style="width: 80px; font-size: 35px;" class="sorting text_center">&nbsp;StNr.&nbsp;</th>';
    t_head += '<th style="width: auto; font-size: 35px;" class="sorting">Paar</th>';
    t_head += '<th style="width: auto; font-size: 35px;" class="sorting">Punkte</th></tr></thead>';
    return t_head;
}

function fix2(wert) {
    var pu=  (Math.round(wert * 100) / 100).toString();
    return pu.replace(".", ",");
}
