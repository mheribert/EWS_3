var ver = 'V3.2017';
var moderator_inhalt = '';

exports.inhalt = function () {
    return moderator_inhalt;
};

exports.mod_seite = function() {
    var HTML_Seite = '<!DOCTYPE html>';

    HTML_Seite += '<head><title>moderator</title><meta http-equiv="expires" content="0">';
    HTML_Seite += '<link rel="stylesheet" href="EWS3.css">' + '\r\n';
    HTML_Seite += '<script src="socket.io/socket.io.js"></script>' + '\r\n';
    HTML_Seite += '<script src="beamod.js"></script></head>' + '\r\n';
    HTML_Seite += '<body><mod_header><table id="kopf_mod" style="width:99%; position:fixed; ">' + '\r\n';
    HTML_Seite += '<tbody><tr><td class="mod_kopf">Runde</td><td class="mod_kopf">WR</td><td class="mod_kopf">Zeitplan</td></tr>' + '\r\n';
    HTML_Seite += '</tbody></table></mod_header><form name="Formular" style="padding: 7vw 0 0 0;"><center><table border="1" rules="rows" width="100%">' + '\r\n';
    HTML_Seite += '<tr><td colspan="8"><table id="mod_inhalt" width="100%"></table></td></tr>' + '\r\n';
    HTML_Seite += '</table></center></form></body></html>' + '\r\n';
    return HTML_Seite;
};

exports.runde = function (io, runden_info, runde) {
    // Rundeninfo
    var HTML_Inhalt = "";
    var rd_ind = 0;
    if (typeof runden_info[0] !== "undefined" && runden_info[0].Tanzrunde_MAX >= runde) {
        for (i = 0; i < runden_info.length; i++) {
            if (runden_info[i].Rundennummer < runde) {
                rd_ind++;
            }
        }

        HTML_Inhalt = '<tr><td class="mod_m" style="background-color:#ffa">' + runden_info[rd_ind].Rundennummer + '/' + runden_info[rd_ind].Tanzrunde_MAX + '</td>';
        HTML_Inhalt += '<td class="mod_m" colspan ="3" style="background-color:#ffa">' + runden_info[rd_ind].Tanzrunde_Text + '</td></tr>';
        if (typeof runden_info[rd_ind].berechnet !== "undefined") {
            HTML_Inhalt += add_platz(runden_info, rd_ind);
        }
        HTML_Inhalt += make_paar(runden_info, rd_ind, 'mod_m');
        if (runden_info[rd_ind].PpR === 2) {
            if (typeof runden_info[rd_ind].berechnet !== "undefined") {
                HTML_Inhalt += add_platz(runden_info, rd_ind + 1);
            }
            HTML_Inhalt += make_paar(runden_info, rd_ind + 1, 'mod_m');
        }
        HTML_Inhalt += '<tr><td class="wr_status" id="content1" colspan="4" align="center"></td></tr>';
        rd_ind += runden_info[rd_ind].PpR;
        if (typeof runden_info[rd_ind] !== "undefined") {
            HTML_Inhalt += '<tr><td class="mod_m" colspan ="3" style="background-color:#ffa">' + runden_info[rd_ind].Rundennummer + '/' + runden_info[rd_ind].Tanzrunde_MAX + '</td></tr>';
            HTML_Inhalt += make_paar(runden_info, rd_ind, 'mod_mk');
            if (runden_info[rd_ind].PpR === 2) {
                HTML_Inhalt += make_paar(runden_info, rd_ind + 1, 'mod_mk');
            }
        }
    }
    moderator_inhalt = HTML_Inhalt;
    io.emit('chat', { msg: 'mod_inhalt', text: HTML_Inhalt });
};

function make_paar(runden_info, rd_ind, cl) {
    var HTML_paar = '<tr><td width="10%" class="' + cl + '">' + runden_info[rd_ind].Startnr + '</td>';
    if (runden_info[rd_ind].Name_Team === null) {        // Einzel
        HTML_paar += '<td width="85" class="' + cl + '"><b>' + runden_info[rd_ind].Dame + '<br>' + runden_info[rd_ind].Herr + '</b><br>' + runden_info[rd_ind].Verein_Name + '</td>';
    } else {                                    // Formationen
        HTML_paar += '<td width="85" class="' + cl + '"><b>' + runden_info[rd_ind].Name_Team + '</b><br>' + runden_info[rd_ind].Verein_Name + '</td>';
    }
    return HTML_paar + '</tr>';
}

function add_platz(runden_info, rd_ind) {
    if (!runden_info[0].ranking_anzeige) { return ''; }
    var platz = '<tr><td colspan="2" class="mod_m">Platz ' + runden_info[rd_ind].Platz + ' mit ';
    if (runden_info[0].ersteRunde !== null) {
        platz += '&nbsp;&nbsp;&nbsp;' + fix2(runden_info[rd_ind].ersteRunde) + '&nbsp;+&nbsp;' + fix2(runden_info[rd_ind].Punkte) + ' = ';
    }
    platz += fix2(runden_info[rd_ind].ersteRunde + runden_info[rd_ind].Punkte);
    return platz + ' Punkten';
}

exports.zeitplan = function (io, connection, ab_rtid) {
    connection
        .query('SELECT RT.RT_ID, RT.Turniernr, RT.Rundenreihenfolge, RT.getanzt, RT.WB, Startklasse.Startklasse, Startklasse_text, Rundentext, Format([Startzeit],"Short Time") AS Zeit, (SELECT Count(PR_ID) AS Ausdr1 FROM Paare_Rundenqualifikation WHERE Paare_Rundenqualifikation.rt_id=rt.rt_id;) AS AnzPaare FROM Tanz_Runden INNER JOIN (Rundentab AS RT LEFT JOIN Startklasse ON RT.Startklasse = Startklasse.Startklasse) ON Tanz_Runden.Runde = RT.Runde WHERE (((RT.Rundenreihenfolge)<999)) ORDER BY RT.Rundenreihenfolge;')
        .on('done', function (data) {
            // Kopf Text
            var beginn = false;
            var HTML_Kopf = 'Zeitplan';
            // Rundeninfo
            var HTML_Inhalt = "";
            for (var i in data) {
                if (typeof ab_rtid === "undefined") {
                    beginn = true;
                } else {
                    if (data[i].RT_ID === ab_rtid) {
                        beginn = true;
                    }
                }
                if (beginn === true && data[i].Rundentext !== "Letzte Startkartenabgabe") {
                    HTML_Inhalt += '<tr><td class="mod_z" width=18%>' + data[i].Zeit + '</td>';
                    if (data[i].Rundentext === "Vorstellung der Tanzpaare" || data[i].Rundentext === "Vorstellung der Formationen") {
                        HTML_Inhalt += '<td class="mod_nb" width=80% id="rt0">';
                    } else if (data[i].Rundentext === "Siegerehrung" && data[i].WB > 0) {
                        HTML_Inhalt += '<td class="mod_ns" width=80% id="rt' + data[i].RT_ID + '" rtid="' + data[i].WB + '">';
                    } else if (data[i].AnzPaare > 0) {
                        HTML_Inhalt += '<td class="mod_nb" width=80% id="rt' + data[i].RT_ID + '">';
                    } else {
                        HTML_Inhalt += '<td class="mod_n" width=80%>';
                    }
                    HTML_Inhalt += data[i].Rundentext + ' ' + (data[i].Startklasse_text || "") + '</td ><td>&nbsp;</td></tr >';
                }
            }
            moderator_inhalt = HTML_Inhalt;
            io.emit('chat', { msg: 'mod_inhalt', text: HTML_Inhalt });
        });
};

exports.wr = function (io, wertungsrichter) {
    var HTML_Inhalt = "";
    for (var i in wertungsrichter) {
        if (wertungsrichter[i].WR_func !== "" && wertungsrichter[i].WR_func !== null) {
            HTML_Inhalt += '<tr><td class="mod_n" width=15%>' + wertungsrichter[i].WR_Kuerzel + '</td>';
            HTML_Inhalt += '<td class="mod_n" width=65%>' + wertungsrichter[i].WR_Vorname + ' ' + wertungsrichter[i].WR_Nachname + '</td >';
            HTML_Inhalt += '<td class="mod_n" width=20%>' + wertungsrichter[i].WR_func + '</td></tr>';
        }
    }
    moderator_inhalt = HTML_Inhalt;
    io.emit('chat', { msg: 'mod_inhalt', text: HTML_Inhalt });
};

exports.vorstellung = function (io, connection, st_kl) {
    var HTML_Inhalt = '';
    var filter;
    if (isNaN(st_kl) ) {
        filter  = 'SELECT Verein_Name, Startnr, [Da_Vorname] & " " & [Da_Nachname] AS DName, [He_Vorname] & " " & [He_Nachname] AS HName, Name_Team, Verein_nr, Paare.Startkl, Startklasse.Startklasse_text ';
        filter += 'FROM Paare INNER JOIN Startklasse ON Paare.Startkl = Startklasse.Startklasse ';
        filter += 'WHERE(Paare.Startkl = "' + st_kl + '" AND Paare.Turniernr = 1 AND Paare.Anwesent_Status = 1) ';
        filter += 'ORDER BY Verein_Name, Startnr;'
    } else {
        st_kl = parseInt(st_kl);
        if (st_kl === 0) {
            filter  = 'SELECT Paare.Verein_Name, Startnr, [Da_Vorname] & " " & [Da_Nachname] AS DName, [He_Vorname] & " " & [He_Nachname] AS HName, Name_Team, Verein_nr, Paare.Startkl, Startklasse.Startklasse_text ';
            filter += 'FROM Paare INNER JOIN Startklasse ON Paare.Startkl = Startklasse.Startklasse ';
            filter += 'WHERE (Turniernr = 1 AND Anwesent_Status = 1) ';
            filter += 'ORDER BY Paare.Verein_Name, Paare.Startnr;';
        } else {
            filter  = 'SELECT Paare.Verein_Name, Startnr, [Da_Vorname] & " " & [Da_Nachname] AS DName, [He_Vorname] & " " & [He_Nachname] AS HName, Name_Team, Verein_nr, Paare.Startkl, Startklasse.Startklasse_text, RT_ID ';
            filter += 'FROM (Paare INNER JOIN  Startklasse ON Paare.Startkl = Startklasse.Startklasse) INNER JOIN Paare_Rundenqualifikation ON Paare.TP_ID = Paare_Rundenqualifikation.TP_ID ';
            filter += 'WHERE (Turniernr=1 AND Anwesent_Status = 1 AND RT_ID=' + st_kl + ') ';
            filter += 'ORDER BY Paare.Verein_Name, Paare.Startnr;';
        }
    }
    connection
        .query(filter)
        .on('done', function (data) {
            var vText = "";

            for (var v in data) {
                if (vText !== data[v].Verein_Name) {
                    HTML_Inhalt += '<tr><td height="10px"></td></tr>' + '\r\n';
                    HTML_Inhalt += '<tr><td colspan="3" class="mod_verein">' + data[v].Verein_Name + '</td ></tr >';
                    vText = data[v].Verein_Name;
                }
                HTML_Inhalt += '<tr height="40" class="mod_paar"><td width=8% align="center">' + data[v].Startnr + '</td><td width=60%>';
                if (data[v].Name_Team === null) {
                    HTML_Inhalt += data[v].DName + " - " + data[v].HName + '</td>';
                } else {
                    HTML_Inhalt += data[v].Name_Team + '</td>';
                }
                HTML_Inhalt += '<td width=32%>' + data[v].Startklasse_text + '</td></tr>' + '\r\n';
            }
            moderator_inhalt = HTML_Inhalt;
            io.emit('chat', { msg: 'mod_inhalt', text: HTML_Inhalt });
        });
};

exports.siegerehrung = function (io, connection, rt_id) {
    connection
        .query('SELECT * FROM View_Rundenablauf WHERE RT_ID =' + rt_id + ' ORDER BY Platz DESC, Startnr;')
        .on('done', function (data) {
            var HTML_Inhalt = make_thead() + '<tbody>';
            var cl = '';
            for (var p in data) {
                HTML_Inhalt += '<tr id="' + data[p].RT_ID + '" class="weiter"  ' + cl + '><td class="mod_s">' + data[p].Platz + '&nbsp;</td>';
                if (data[p].Name_Team === null) {
                    HTML_Inhalt += '<td class="mod_s">' + data[p].Startnr + '</td><td  class="mod_s">' + data[p].Dame + ' - ' + data[p].Herr + '</td>';
                } else {
                    HTML_Inhalt += '<td class="mod_s">' + data[p].Startnr + '</td><td class="text_left">' + data[p].Name_Team + '</td>';
                }
                data[p].jetztRunde = data[p].jetztRunde || 0;
                punkte = data[p].jetztRunde.toFixed(2);
                HTML_Inhalt += '<td class="mod_pkte">' + punkte + '</td></tr>';
            }
            HTML_Inhalt += '</tbody>';

            moderator_inhalt = HTML_Inhalt;
            io.emit('chat', { msg: 'mod_inhalt', text: HTML_Inhalt });
        });

};

function make_thead() {
    var t_head = '<thead><tr ><th class="mod_s">Platz</th>';
    t_head += '<th class="mod_s">&nbsp;StNr.&nbsp;</th>';
    t_head += '<th class="mod_s">Paar</th>';
    t_head += '<th class="mod_s">Punkte</th></tr></thead>';
    return t_head;
}

function fix2(wert) {
    var pu = (Math.round(wert * 100) / 100).toString();
    return pu.replace(".", ",");
}
