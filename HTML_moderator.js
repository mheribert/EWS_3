var ver              = 'V3.2.00';
var moderator_inhalt = '';

exports.inhalt = function () {
    return moderator_inhalt;
};

exports.mod_seite = function() {
    var HTML_Seite = '<!DOCTYPE html>';

    HTML_Seite += '<head><title>moderator</title><meta http-equiv="expires" content="0">';
    HTML_Seite += '<link rel="stylesheet" href="EWS3.css">' + '\r\n';
    HTML_Seite += '<script src="socket.io/socket.io.js"></script>' + '\r\n';
    HTML_Seite += '<script src="EWS3.js" ></script></head>' + '\r\n';

    HTML_Seite += '<body><form name="Formular"><center><table border="1" rules="rows" width="100%">' + '\r\n';
    HTML_Seite += '<tr><td colspan="8"><table id="kopf_mod" style="width:100%; table-layout:fixed; ">' + '\r\n';
    HTML_Seite += '<tr><td class="mod_kopf">Runde</td><td class="mod_kopf">WR</td><td class="mod_kopf">Zeitplan</td></tr>' + '\r\n';
    HTML_Seite += '</table></td></tr>' + '\r\n';
    HTML_Seite += '<tr><td colspan="8"><table id="mod_inhalt" width="100%"></table></td></tr>' + '\r\n';
    HTML_Seite += '</table></center></form></body></html>' + '\r\n';
    return HTML_Seite;
};

exports.runde = function(io, runden_info, runde) {
    // Rundeninfo
    var HTML_Inhalt = "";
    var rd_ind = 0;
    if (typeof runden_info[0] !== "undefined" && runden_info[0].Tanzrunde_MAX >= runde) {
        for (var i = 1; i < runde; i++) {
            rd_ind += parseInt(runden_info[i].PpR);
        }

        HTML_Inhalt = '<tr><td class="mod_m" style="background-color:#ff8">' + runden_info[rd_ind].Rundennummer + '/' + runden_info[rd_ind].Tanzrunde_MAX + '</td>';
        HTML_Inhalt += '<td class="mod_m" colspan ="3" style="background-color:#ff8">' + runden_info[rd_ind].Tanzrunde_Text + '</td></tr>';
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
            HTML_Inhalt += '<tr><td class="mod_m" colspan ="3" style="background-color:#ff8">' + runden_info[rd_ind].Rundennummer + '/' + runden_info[rd_ind].Tanzrunde_MAX + '</td></tr>';
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
        .query('SELECT RT.RT_ID, RT.Turniernr, RT.Rundenreihenfolge, Startklasse.Startklasse, Startklasse_text, Rundentext, Format([Startzeit],"Short Time") AS Zeit, (SELECT Count(PR_ID) AS Ausdr1 FROM Paare_Rundenqualifikation WHERE Paare_Rundenqualifikation.rt_id=rt.rt_id;) AS AnzPaare FROM Tanz_Runden INNER JOIN (Rundentab AS RT LEFT JOIN Startklasse ON RT.Startklasse = Startklasse.Startklasse) ON Tanz_Runden.Runde = RT.Runde WHERE (((RT.Rundenreihenfolge)<999)) ORDER BY RT.Rundenreihenfolge;')
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
                    if (data[i].Rundentext === "Vorstellung der Tanzpaare") {
                        HTML_Inhalt += '<td class="mod_nb" width=80% id="rt0">';
                    } else if (data[i].Rundentext === "Siegerehrung") {
                        HTML_Inhalt += '<td class="mod_n" width=80% id="rt' + data[i].RT_ID + '">';
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
    if (st_kl === "0") {
        filter = 'SELECT Verein_Name, Startnr, [Da_Vorname] & " " & [Da_Nachname] AS DName, [He_Vorname] & " " & [He_Nachname] AS HName, Name_Team, Verein_nr, Paare.Startkl, Startklasse.Startklasse_text FROM Paare INNER JOIN Startklasse ON Paare.Startkl = Startklasse.Startklasse WHERE (Turniernr = 1 AND Anwesent_Status = 1) ORDER BY Paare.Verein_Name, Paare.Startnr;'
    } else {
        filter = 'SELECT Verein_Name, Startnr, [Da_Vorname] & " " & [Da_Nachname] AS DName, [He_Vorname] & " " & [He_Nachname] AS HName, Name_Team, Verein_nr, Paare.Startkl, Startklasse.Startklasse_text FROM Paare INNER JOIN Startklasse ON Paare.Startkl = Startklasse.Startklasse WHERE (Paare.Startkl ="' + st_kl + '" AND Paare.Turniernr = 1 AND Paare.Anwesent_Status = 1) ORDER BY Verein_Name, Startnr;'
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
                HTML_Inhalt += '<tr height="40" class="mod_paar"><td width=10% align="center">' + data[v].Startnr + '</td><td width=70%>';
                if (data[v].Name_Team === null) {
                    HTML_Inhalt += data[v].DName + " - " + data[v].HName + '</td>';
                } else {
                    HTML_Inhalt += data[v].Name_Team + '</td>';
                }
                HTML_Inhalt += '<td width=20%>' + data[v].Startklasse_text + '</td></tr>' + '\r\n';
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
                punkte = fix2(data[p].jetztRunde);
                HTML_Inhalt += '<td class="mod_s">' + punkte + '</td></tr>';
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
