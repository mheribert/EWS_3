var ver = 'V3.2020';
const rei = false;

exports.wr_login = function (wertungsrichter, title) {
    var HTML_Seite = '<!DOCTYPE html>';
    HTML_Seite += '<head><title>login</title><meta http-equiv="expires" content="0">';

    HTML_Seite += '<link rel="stylesheet" href="EWS3.css">' + '\r\n';
    HTML_Seite += '<script> window.onload = start;' + '\r\n';
    HTML_Seite += 'function start() {t = document.getElementsByClassName("wr_l"); for (var i = 0; i < t.length; i++) { t[i].setAttribute("onclick", "wr_onclick(event)"); }';
    HTML_Seite += 't = document.getElementsByClassName("wr_prl"); for (var i = 0; i < t.length; i++) { t[i].setAttribute("onclick", "wr_onclick(event)"); } }' + '\r\n';
    HTML_Seite += 'function wr_onclick(e) { e = e || window.event; var tar = e.target || e.srcElement; var eingabe = window.prompt(tar.innerHTML + " bitte geben Sie das Passwort ein", "");';
    HTML_Seite += 'if (eingabe != null) { document.getElementById("wr_id").value = tar.attributes.max.value; document.getElementById("passwort").value = eingabe; document.forms["Login"].submit(); } }';
    HTML_Seite += '</script></head>';

    HTML_Seite += '<body><form name="Login" action=/login method=post><center><table border="1" rules="rows">' + '\r\n';
    HTML_Seite += '<tr><td class="ind_o" colspan="2">' + title + '<input type="hidden" name="wr_id" id="wr_id"><input type="hidden" name="passwort" id="passwort"></td></tr>' + '\r\n';
    for (var i in wertungsrichter) {
        if (wertungsrichter[i].WR_Azubi === true) {
            HTML_Seite += '<tr><td class="wr_prm">' + wertungsrichter[i].WR_Kuerzel + '</td><td class="wr_prl" max="' + wertungsrichter[i].WR_ID + '">' + wertungsrichter[i].WR_Vorname + ' ' + wertungsrichter[i].WR_Nachname + '</td></tr>' + '\r\n';
        } else {
            HTML_Seite += '<tr><td class="wr_m">' + wertungsrichter[i].WR_Kuerzel + '</td><td class="wr_l" max="' + wertungsrichter[i].WR_ID + '">' + wertungsrichter[i].WR_Vorname + ' ' + wertungsrichter[i].WR_Nachname + '</td></tr>' + '\r\n';
        }
    }
    HTML_Seite += '</table></center></form></body></html>';
    return HTML_Seite;
};

exports.blankPage = function (rd_ind, wr_name, wr_id, runden_info, res) {
    var HTML_Seite = make_HTMLhead(wr_id, runden_info, "judgetool") + '\r\n';
    HTML_Seite += make_kopf(rd_ind, runden_info, 0, wr_name) + '\r\n';
    HTML_Seite += '<tr id="anzeige_body">' + '\r\n';
    HTML_Seite += '<td class="main" height="450px"></td></tr>';
    HTML_Seite += make_absenden(false, false) + '</table></center></form></body></html>';
    res.send(HTML_Seite);
};

exports.mkPage = function (rd_ind, wr_name, wr_id, runden_info, res, func) {
    var HTML_Seite = make_HTMLhead(wr_id, runden_info, "mehrkampf") + '\r\n';
    HTML_Seite += make_kopf(rd_ind, runden_info, 0, wr_name) + '\r\n';
    HTML_Seite += '<script> const wr_func="' + func + '" </script>';
    HTML_Seite += '<tr><td class="mk_sel" align="center"><table width="100%" align="center">';
    HTML_Seite += '<tr><td width="50%"><label>Station:<select class="select" name="station" id="station"><option>---</option></select></label></td>';
    HTML_Seite += '<td width="50%"><label id="stkl_reload">Startklasse:<select class="select" name="klasse" id="klasse"><option>---</option></select></label></td></tr>';
    HTML_Seite += '</table></td></tr>';
    HTML_Seite += '<tr id="wertungen">' + '\r\n';
    HTML_Seite += '<td class="main" height="300px"></td></tr>';
    HTML_Seite += make_absenden(false, false) + '</table></center></form></body></html>';
    res.send(HTML_Seite);
};

exports.wait = function (rd_ind, runden_info, text, wr_name, wr_id, io) {
    var HTML_Seite = make_kopf(rd_ind, runden_info, 0, wr_name) + '\r\n';
    HTML_Seite += '<tr><td height="400px" align="center"><div class="wr_status" id="content1">' + text + '</div></td></tr>';
    HTML_Seite += make_absenden(false, false);
    HTML_Seite += '</table></center></form>';
    io.sockets.emit('chat', { msg: 'body', WR: wr_id, HTML: HTML_Seite });
};

exports.BS_Seite = function (rd_ind, runden_info, wr_name, wr_id, tausch, io) {
    var st_kl = runden_info[0].Startklasse;
    var sei;

    var HTML_Seite = make_kopf(rd_ind, runden_info, 2, wr_name, tausch) + '\r\n';
    HTML_Seite += '<tr id="anzeige_body">' + '\r\n';
    for (var s = 1; s <= runden_info[rd_ind].PpR; s++) {
        sei = s;
        if (tausch === true && runden_info[rd_ind].PpR === 2) { sei = 3 - s; }
        HTML_Seite += '<td class="main" height="400px"><table align="center" border="0" cellpadding="0" cellspacing="0">' + '\r\n';
        HTML_Seite += '<tr><td class="akro" colspan = "21">Tanztechnik, Choreographie, Tanzfiguren, Akrobatikfiguren</td></tr>';
        HTML_Seite += '<tr><td height="10px"></td></tr>';
        HTML_Seite += '<tr id="gs' + sei + '" class="kriterium">';
        for (t = 0; t < 21; t++) {
            if (t % 2 === 0) {
                HTML_Seite += '<td class="bs_wert">' + t / 2 + '</td>';
            } else {
                HTML_Seite += '<td class="bs_wert">-</td>';
            }
        }
        HTML_Seite += '<td><input class="punkte" id="wgs' + sei + '" value="" type="hidden" name="wgs' + sei + '" max="10"></td></tr>' + '\r\n';
        HTML_Seite += '</table></td>' + '\r\n';
    }
    HTML_Seite += make_absenden(true, false, tausch === true && runden_info[rd_ind].PpR === 2) + '</table></center></form>';

    io.sockets.emit('chat', { msg: 'body', WR: wr_id, HTML: HTML_Seite, ausw: 'BS_' });
};

exports.BS_BY_BWSeite = function (rd_ind, runden_info, wr_name, wr_id, tausch, io) {
    var st_kl = runden_info[0].Startklasse;
    var trunde = runden_info[0].RundeArt;
    var sei;

    HTML_Seite = make_kopf(rd_ind, runden_info, 2, wr_name, tausch) + '\r\n';
    HTML_Seite += '<tr id="anzeige_body">' + '\r\n';
    for (var s = 1; s <= runden_info[rd_ind].PpR; s++) {
        sei = s;
        if (tausch === true && runden_info[rd_ind].PpR === 2) { sei = 3 - s; }
        HTML_Seite += '<td align="center" id="couple' + s + '"><table align="center" border="0" cellpadding="0" cellspacing="0">' + '\r\n';
        HTML_Seite += make_bs_inp('gs' + sei, 10, 'Grundschritt (Rhythmus & Fu&szlig;technik)', st_kl) + '\r\n';
        HTML_Seite += make_bs_inp('bd' + sei, 10, 'Basic Dancing, Lead & Follow, Harmonie', st_kl) + '\r\n';
        HTML_Seite += make_bs_inp('tf' + sei, 10, 'Tanzfiguren (einfache, highlight)', st_kl) + '\r\n';
        HTML_Seite += make_bs_inp('in' + sei, 10, 'Interpretation (Figuren, spontane Interpretation)', st_kl) + '\r\n';
        HTML_Seite += '<tr><td height="30"></td></tr ></table></td>' + '\r\n';
    }
    HTML_Seite += '</tr>' + make_absenden(true, false, tausch === true && runden_info[rd_ind].PpR === 2) + '</table></center></form>';

    io.sockets.emit('chat', { msg: 'body', WR: wr_id, HTML: HTML_Seite, ausw: 'BS_' });
};

exports.MK_WB_Seite = function (rd_ind, runden_info, wr_name, wr_id, tausch, io, wr_func) {
    var st_kl = runden_info[0].Startklasse;
    var trunde = runden_info[0].Runde.substr(0, 4);
    var ausw = "BS_";
    var sei;

    HTML_Seite = make_kopf(rd_ind, runden_info, 2, wr_name, tausch) + '\r\n';
    HTML_Seite += '<tr id="anzeige_body">' + '\r\n';
    for (var s = 1; s <= runden_info[rd_ind].PpR; s++) {
        sei = s;
        if (tausch === true && runden_info[rd_ind].PpR === 2) { sei = 3 - s; }
        HTML_Seite += '<td align="center" id="couple' + sei + '"><table align="center" border="0" cellpadding="0" cellspacing="0">' + '\r\n';
        if (wr_func === "MA") {
            if (trunde === "MK_5") {
                HTML_Seite += make_inpMK('mk_th' + sei, 10, 'Figuren', st_kl) + '\r\n';
                HTML_Seite += '<tr><td height="230"></td></tr>' + '\r\n';
            } else {
                if (trunde === "MK_3" || trunde === "MK_4") {
                    HTML_Seite += make_inpMK('mk_td' + sei, 10, 'Dame', "Ob") + '\r\n';
                    HTML_Seite += '<tr><td height="60"></td></tr>' + '\r\n';
                    HTML_Seite += make_inpMK('mk_th' + sei, 10, 'Herr', "Ob") + '\r\n';
                    HTML_Seite += '<tr><td height="100"></td></tr>' + '\r\n';
                } else {
                    HTML_Seite += '<tr><td height="50"></td></tr>' + '\r\n';
                    HTML_Seite += make_inpMKText('mk_td' + sei, 0, "Dame", st_kl) + '\r\n';
                    HTML_Seite += '<tr><td height="100"></td></tr>' + '\r\n';
                    HTML_Seite += make_inpMKText('mk_th' + sei, 0, "Herr", st_kl) + '\r\n';
                    HTML_Seite += '<tr><td height="30"></td></tr>' + '\r\n';
                    ausw = "MK_T";
               }
            }
        } else {
            if (trunde === "End_" || trunde === "MK_3" || trunde === "MK_4" || trunde === "MK_5") {
                HTML_Seite += make_inpMK('mk_td' + sei, 7, 'Dame Technik & Haltung', st_kl) + '\r\n';
                HTML_Seite += make_inpMK('mk_dd' + sei, 3, 'Dame Dynamik & Takt', st_kl) + '\r\n';
                HTML_Seite += '<tr><td height="50"></td></tr>' + '\r\n';
                HTML_Seite += make_inpMK('mk_th' + sei, 7, 'Herr Technik & Haltung', st_kl) + '\r\n';
                HTML_Seite += make_inpMK('mk_dh' + sei, 3, 'Herr Dynamik & Takt', st_kl) + '\r\n';
           } else {
                HTML_Seite += '<tr><td height="270">Kein Einsatz</td></tr>' + '\r\n';
            }
            ausw = "MK_B"
        }
        HTML_Seite += '<tr><td height="30"></td></tr ></table></td>' + '\r\n';
    }
    HTML_Seite += '</tr>' + make_absenden(true, false, tausch === true && runden_info[rd_ind].PpR === 2) + '</table></center></form>';

    io.sockets.emit('chat', { msg: 'body', WR: wr_id, HTML: HTML_Seite, ausw: ausw });
};

exports.BS_BW_BWSeite = function (rd_ind, runden_info, wr_name, wr_id, tausch, io) {
    var st_kl = runden_info[0].Startklasse;
    var trunde = runden_info[0].RundeArt;
    var sei;

    HTML_Seite = make_kopf(rd_ind, runden_info, 2, wr_name, tausch) + '\r\n';
    HTML_Seite += '<tr id="anzeige_body">' + '\r\n';
    for (var s = 1; s <= runden_info[rd_ind].PpR; s++) {
        sei = s;
        if (tausch === true && runden_info[rd_ind].PpR === 2) { sei = 3 - s; }
        HTML_Seite += '<td align="center" id="couple' + s + '"><table align="center" border="0" cellpadding="0" cellspacing="0">' + '\r\n';
        if (st_kl.substr(0, 6) === "BS_BW_") {
            HTML_Seite += make_bs_inp('th' + sei, 10, 'Technik Herr', st_kl) + '\r\n';
        } else {
            HTML_Seite += make_bs_inp('th' + sei, 10, 'Technik', st_kl) + '\r\n';
        }
        if (st_kl.substr(0, 6) === "BS_BW_") {
            HTML_Seite += make_bs_inp('td' + sei, 10, 'Technik Dame', st_kl) + '\r\n';
        } else {
            HTML_Seite += '<tr><td class="bs_ersatz"><input name="wtd' + sei + '" id="wtd' + sei + '" type="hidden" value="0"></td></tr>';
        }
        HTML_Seite += make_bs_inp('ta' + sei, 10, 'Tanz', st_kl) + '\r\n';
        HTML_Seite += '<tr><td colspan="21"><hr></td></tr>' + '\r\n';
        if (st_kl === "BS_BW_FO" || st_kl === "BS_BW_HA") {
            HTML_Seite += make_bs_inp('ak' + sei, 10, 'Akrobatik', st_kl) + '\r\n';
        } else {
            HTML_Seite += '<tr><td class="bs_ersatz"><input name="wak' + sei + '" id="wak' + sei + '" type="hidden" value="0"></td></tr>';
        }
        HTML_Seite += '<tr><td class="bs_schmal"></td></tr>';
        HTML_Seite += '<tr><td colspan="21"><hr></td></tr>' + '\r\n';
        HTML_Seite += make_bs_feh(sei, 1) + '</table></td> ' + '\r\n';
    }
    HTML_Seite += '</tr>' + make_absenden(true, false, tausch === true && runden_info[rd_ind].PpR === 2) + '</table></center></form>';

    io.sockets.emit('chat', { msg: 'body', WR: wr_id, HTML: HTML_Seite, ausw: 'BS_' });
};

exports.BS_HE_Seite = function (rd_ind, runden_info, wr_name, wr_id, tausch, io) {
    var st_kl = runden_info[0].Startklasse;
    var trunde = runden_info[0].RundeArt;
    var sei;

    HTML_Seite = make_kopf(rd_ind, runden_info, 2, wr_name, tausch) + '\r\n';
    HTML_Seite += '<tr id="anzeige_body">' + '\r\n';
    for (var s = 1; s <= runden_info[rd_ind].PpR; s++) {
        sei = s;
        if (tausch === true && runden_info[rd_ind].PpR === 2) { sei = 3 - s; }
        HTML_Seite += '<td align="center" id="couple' + s + '"><table align="center" border="0" cellpadding="0" cellspacing="0">' + '\r\n';
        HTML_Seite += make_bs_inp('th' + sei, 5, 'Technik Herr', st_kl) + '\r\n';
        HTML_Seite += make_bs_inp('td' + sei, 5, 'Technik Dame', st_kl) + '\r\n';
        HTML_Seite += make_bs_inp('ta' + sei, 5, 'Tanzfiguren', st_kl) + '\r\n';
        HTML_Seite += make_bs_inp('ak' + sei, 5, 'Choreographie', st_kl) + '\r\n';
        HTML_Seite += '<tr><td class="bs_schmal"></td></tr>';
        HTML_Seite += '<tr><td colspan="21"><hr></td></tr>' + '\r\n';
        HTML_Seite += make_bs_feh(sei, 1) + '</table></td> ' + '\r\n';
    }
    HTML_Seite += '</tr>' + make_absenden(true, false, tausch === true && runden_info[rd_ind].PpR === 2) + '</table></center></form>';

    io.sockets.emit('chat', { msg: 'body', WR: wr_id, HTML: HTML_Seite, ausw: 'BS_' });
};

exports.BS_SL_Seite = function (rd_ind, runden_info, wr_name, wr_id, tausch, io) {
    var st_kl = runden_info[0].Startklasse;
    var trunde = runden_info[0].RundeArt;
    var sei;

    HTML_Seite = make_kopf(rd_ind, runden_info, 2, wr_name, tausch) + '\r\n';
    HTML_Seite += '<tr id="anzeige_body">' + '\r\n';
    for (var s = 1; s <= runden_info[rd_ind].PpR; s++) {
        sei = s;
        if (tausch === true && runden_info[rd_ind].PpR === 2) { sei = 3 - s; }
        HTML_Seite += '<td align="center" id="couple' + s + '"><table align="center" border="0" cellpadding="0" cellspacing="0">' + '\r\n';
        HTML_Seite += '<tr><td class="bs_schmal"></td></tr>';
        HTML_Seite += make_bs_inp('th' + sei, 10, 'Takt, Rhythmus', st_kl) + '\r\n';
        HTML_Seite += '<tr><td class="bs_schmal"></td></tr>';
        HTML_Seite += make_bs_inp('td' + sei, 10, 'Bewegung, Balancen', st_kl) + '\r\n';
        HTML_Seite += '<tr><td class="bs_schmal"></td></tr>';
        HTML_Seite += make_bs_inp('ta' + sei, 10, 'Präsentation, Charakteristik, Synchronität', st_kl) + '\r\n';
        HTML_Seite += '<tr><td class="bs_schmal"></td></tr>';
        HTML_Seite += '<tr><td colspan="21"><hr></td></tr>' + '\r\n';
        HTML_Seite += make_bs_feh(sei, 0) + '</table></td> ' + '\r\n';
    }
    HTML_Seite += '</tr>' + make_absenden(true, false, tausch === true && runden_info[rd_ind].PpR === 2) + '</table></center></form>';

    io.sockets.emit('chat', { msg: 'body', WR: wr_id, HTML: HTML_Seite, ausw: 'BS_' });
};

exports.BS_BY_Observer = function (rd_ind, runden_info, wr_name, wr_id, io) {
    var HTML_Seite = make_kopf(rd_ind, runden_info, 2, wr_name, false) + '\r\n';
    HTML_Seite += '<tr id="anzeige_body">' + '\r\n';

    for (var s = 1; s <= runden_info[rd_ind].PpR; s++) {
        HTML_Seite += '<td align="center" id="couple' + s + '" width="450px"><table border="0" cellpadding="1" cellspacing="5">';
        HTML_Seite += '<tr id="seite' + s + '">' + '\r\n';
        HTML_Seite += '<td height="140px"></td></tr>';
        HTML_Seite += '<tr><td class="bs_verw leer" onclick="verwarnung(event)" value="0">Verwarnung<input name="wverw' + s + '" type="hidden"></td></td></tr>' + '\r\n';
        HTML_Seite += '<td height="140px"></td></tr>';
        HTML_Seite += '</table></td>' + '\r\n';
    }
    HTML_Seite += '</tr>' + make_absenden(true, true, false) + '</table></center></form>';

    io.sockets.emit('chat', { msg: 'body', WR: wr_id, HTML: HTML_Seite, ausw: 'OB_' });
};

exports.BW_Seite = function (rd_ind, runden_info, wr_name, wr_id, tausch, io) {
    var st_kl = runden_info[0].Startklasse;
    var trunde = runden_info[0].RundeArt;
    var sei;

    HTML_Seite = make_kopf(rd_ind, runden_info, 2, wr_name, tausch) + '\r\n';
    HTML_Seite += '<tr id="anzeige_body">' + '\r\n';
    for (var s = 1; s <= runden_info[rd_ind].PpR; s++) {
        sei = s;
        if (tausch === true && runden_info[rd_ind].PpR === 2) { sei = 3 - s; }
        HTML_Seite += '<td align="center" id="couple' + s + '">' + '\r\n';
        if (trunde === 'ER') {
            HTML_Seite += make_inpBW('gs' + sei, 10, 'Grundschritt (Rhythmus & Fu&szlig;technik)', st_kl) + '\r\n';
            HTML_Seite += make_inpBW('bd' + sei, 10, 'Basic Dancing, Lead & Follow, Harmonie', st_kl) + '\r\n';
            HTML_Seite += make_inpBW('tf' + sei, 10, 'Tanzfiguren (Komplexe, Highlight)', st_kl) + '\r\n';
            HTML_Seite += make_inpBW('in' + sei, 10, 'Interpretation (Komplexe und Highlight Figuren)', st_kl) + '\r\n';
            HTML_Seite += make_inpBW('si' + sei, 10, 'Spontane Interpretation', st_kl) + '\r\n';
            HTML_Seite += make_inpBW('dp' + sei, 10, 'Dance Performance', st_kl) + '\r\n';
        } else {
            HTML_Seite += make_inpBW('gs' + sei, 10, 'Tanztechnik', st_kl) + '\r\n';
            HTML_Seite += make_inpBW('tf' + sei, 10, 'Tanzfiguren', st_kl) + '\r\n';
            HTML_Seite += make_inpBW('in' + sei, 10, 'Interpretation der Musik', st_kl) + '\r\n';
            HTML_Seite += make_inpBW('dp' + sei, 10, 'Dance Performance', st_kl) + '\r\n';
        }
    }
    HTML_Seite += '</tr>';
    HTML_Seite += make_absenden(true, false, tausch === true && runden_info[rd_ind].PpR === 2) + '</table></center></form>';

    io.sockets.emit('chat', { msg: 'body', WR: wr_id, HTML: HTML_Seite, ausw: 'BW_' });
};

exports.BW_NG_Seite = function (rd_ind, runden_info, wr_name, wr_id, tausch, io) {
    var st_kl = runden_info[0].Startklasse;
    var trunde = runden_info[0].RundeArt;
    var sei;

    HTML_Seite = make_kopf(rd_ind, runden_info, 2, wr_name, tausch) + '\r\n';
    HTML_Seite += '<tr>' + '\r\n';
    for (var s = 1; s <= runden_info[rd_ind].PpR; s++) {
        sei = s;
        if (tausch === true && runden_info[rd_ind].PpR === 2) { sei = 3 - s; }
        HTML_Seite += '<td bgcolor = "#dddddd" style = "font-family: Arial; padding-left: 10px; padding-right: 10px; font-size: 16px;"><table align="center" border="0">';
        HTML_Seite += kriterium_text("Schritt Dame", "Schritt Herr");
        HTML_Seite += '<tr><td>' + make_inpNG_BW('ng_ttd' + sei, 5, '', st_kl, '') + '</td>';
        HTML_Seite += '<td colspan="2">' + make_inpNG_BW('ng_tth' + sei, 5, '', st_kl, '') + '</td></tr>' + '\r\n';
        HTML_Seite += kriterium_text("Basic Dancing,Lead & Follow,Harmonie,Dance Performance");
        HTML_Seite += '<tr><td>' + make_inpNG_BW('ng_bda' + sei, 5, '', st_kl, '') + '</td>';
        HTML_Seite += '<td>' + make_inpNG_BW('ng_dap' + sei, 4, '', st_kl, '0') + '</td>';
        HTML_Seite += '<td>' + make_inpNG_BW('ng_bdb' + sei, 1, '', st_kl, '0') + '</td></tr>' + '\r\n';
        HTML_Seite += kriterium_text("Tanzfiguren");
        HTML_Seite += '<tr><td>' + make_inpNG_BW('ng_fta' + sei, 5, '', st_kl, '') + '</td>';
        HTML_Seite += '<td>' + make_inpNG_BW('ng_fts' + sei, 4, '', st_kl, '0') + '</td>';
        HTML_Seite += '<td>' + make_inpNG_BW('ng_ftb' + sei, 1, '', st_kl, '0') + '</td></tr>' + '\r\n';
        HTML_Seite += kriterium_text("Musik Interpretation");
        HTML_Seite += '<tr><td>' + make_inpNG_BW('ng_inf' + sei, 5, '', st_kl, '') + '</td>';
        HTML_Seite += '<td>' + make_inpNG_BW('ng_ins' + sei, 4, '', st_kl, '0') + '</td>';
        HTML_Seite += '<td>' + make_inpNG_BW('ng_inb' + sei, 1, '', st_kl, '0') + '</td></tr>' + '\r\n';
        HTML_Seite += '<tr><td height="30"></td></tr></table></td>';
    }
    HTML_Seite += '</tr>';
    HTML_Seite += make_absenden(true, false, tausch === true && runden_info[rd_ind].PpR === 2) + '</table></center></form>';

    io.sockets.emit('chat', { msg: 'body', WR: wr_id, HTML: HTML_Seite, ausw: 'BW_NG' });
};

exports.BW_Form_Seite = function (rd_ind, runden_info, akrobatiken, wr_func, wr_name, wr_id, io) {
    var st_kl = runden_info[0].Startklasse;
    var trunde = runden_info[0].RundeArt;
    sei = 1;
    var HTML_Seite = make_kopf(rd_ind, runden_info, 1, wr_name) + '\r\n';
    HTML_Seite += '<tr><td bgcolor = "#dddddd" style = "font-family: Arial; padding-left: 10px; padding-right: 10px; font-size: 16px;"><table align="center" border="0">';
    HTML_Seite += kriterium_text("Schritt / Rhythmus", "Synchronität");
    HTML_Seite += '<tr><td>' + make_inpNG_BW('ng_ttd' + sei, 5, '', st_kl, '') + '</td>';
    HTML_Seite += '<td colspan="2">' + make_inpNG_BW('ng_tth' + sei, 5, '', st_kl, '') + '</td></tr>' + '\r\n';
    HTML_Seite += kriterium_text("Tanzfiguren Ausfühtrung, Schwierigkeit/Vielfalt/Originalität, Bonus");
    HTML_Seite += '<tr><td>' + make_inpNG_BW('ng_bda' + sei, 5, '', st_kl, '') + '</td>';
    HTML_Seite += '<td>' + make_inpNG_BW('ng_dap' + sei, 4, '', st_kl, '0') + '</td>';
    HTML_Seite += '<td>' + make_inpNG_BW('ng_bdb' + sei, 1, '', st_kl, '0') + '</td></tr>' + '\r\n';
    HTML_Seite += kriterium_text("Choreografie/Performance/Präsentation, Idee,    Bonus");
    HTML_Seite += '<tr><td>' + make_inpNG_BW('ng_fta' + sei, 5, '', st_kl, '') + '</td>';
    HTML_Seite += '<td>' + make_inpNG_BW('ng_fts' + sei, 4, '', st_kl, '0') + '</td>';
    HTML_Seite += '<td>' + make_inpNG_BW('ng_ftb' + sei, 1, '', st_kl, '0') + '</td></tr>' + '\r\n';
    HTML_Seite += kriterium_text("Synchronität und Ausführung, Positionen,  Bonus");
    HTML_Seite += '<tr><td>' + make_inpNG_BW('ng_inf' + sei, 5, '', st_kl, '') + '</td>';
    HTML_Seite += '<td>' + make_inpNG_BW('ng_ins' + sei, 4, '', st_kl, '0') + '</td>';
    HTML_Seite += '<td>' + make_inpNG_BW('ng_inb' + sei, 1, '', st_kl, '0') + '</td></tr>' + '\r\n';
    HTML_Seite += '<tr><td height="30"></td></tr></table></td></tr>';

    HTML_Seite += make_absenden(true, false) + '</table></center></form>';

    io.sockets.emit('chat', { msg: 'body', WR: wr_id, HTML: HTML_Seite, ausw: 'BW_NG' });
};

exports.BW_Observer = function (rd_ind, runden_info, wr_name, wr_id, io) {
    var krit = ['sbs1', 'Side by Side', '2x8', 'sbs2', 'Side by Side', '4x8', 'akro', 'Acrobatic', '2', 'high', 'Highlight', '4'];
    var HTML_Seite = make_kopf(rd_ind, runden_info, 2, wr_name, false) + '\r\n';
    HTML_Seite += '<tr id="anzeige_body">' + '\r\n';

    for (var s = 1; s <= runden_info[rd_ind].PpR; s++) {
        HTML_Seite += '<td align="center" id="couple' + s + '" width="450px"><table border="0" cellpadding="1" cellspacing="5">';
        if (runden_info[0].Startkl.substring(0, 3) === "BS_") {
            HTML_Seite += '<tr id="seite' + s + '">' + '\r\n';
            HTML_Seite += '<td height="140px"></td></tr>';
            HTML_Seite += '<tr><td class="bs_text">Wenn Runde beendet ist, absenden.</td></tr> ';
            HTML_Seite += '<td height="140px"></td></tr>';
        } else {
            for (var i = 0; i < krit.length; i = i + 3) {
                HTML_Seite += '<tr id="' + krit[i] + s + '"><td class="spalte">-</td><td class="spalte_br"><table width ="100%">' + '\r\n';
                HTML_Seite += '<tr><td colspan="2">' + krit[i + 1] + '</td></tr>' + '\r\n';
                HTML_Seite += '<tr><td id="t' + krit[i] + s + '" style="text-align: left;">' + krit[i + 2] + '</td>';
                HTML_Seite += '<td><input readonly ="" value ="0" name="w' + krit[i] + s + '" id="w' + krit[i] + s + '" style="width: 30px; text-align: center; border-radius:5px;"></td></tr>' + '\r\n';
                HTML_Seite += '</table></td><td class="spalte">+</td></tr>' + '\r\n';
            }
            HTML_Seite += '<tr><td height = "5px"> </td></tr>';
            HTML_Seite += '<tr><td class="verwbutton leer">Acrobatic<input name="wakrovw' + s + '" type="hidden"></td><td class="verwbutton leer">Figuren jun<input name="wjuniorvw' + s + '" type="hidden"></td><td class="verwbutton leer">Kleidung<input name="wkleidungvw' + s + '" type="hidden"></td></tr>' + '\r\n';
            HTML_Seite += '<tr><td class="verwbutton leer">Tanzbereich<input name="wtanzbereichvw' + s + '" type="hidden"></td><td class="verwbutton leer">Tanzzeit<input name="wtanzzeitvw' + s + '" type="hidden"></td><td class="verwbutton leer">2.Aufruf   Verlassen TF   Unsportlich<input name="waufrufvw' + s + '" type="hidden"></td></tr>' + '\r\n';
        }
        HTML_Seite += '</table></td>' + '\r\n';
    }

    HTML_Seite += '</tr>' + make_absenden(true, true, false) + '</table></center></form>';

    io.sockets.emit('chat', { msg: 'body', WR: wr_id, HTML: HTML_Seite, ausw: 'OB_' });
};

exports.BW_ObsCheck = function (rd_ind, wertungsrichter, wertungen, runden_info, runde, wr_name, wr_id, io, new_guidelines) {
    var HTML_Arr = new Object();
    var HTML_Seite;
    var cgi_val;
    var tp_id;
    var trunde = runden_info[0].RundeArt;
    var wr;
    var seite;

    if (runde <= runden_info.length) {
        HTML_Arr[0] = make_kopf(rd_ind, runden_info, 2, wr_name) + '\r\n';
        // headline
        for (seite = 1; seite < runden_info[rd_ind].PpR + 1; seite++) {
            if (runden_info[rd_ind + seite - 1].Name_Team === null) {
                HTML_Seite = '<tr id = "anzeige_body"><td>' + runden_info[rd_ind + seite - 1].Da_Nachname + ' & ' + runden_info[rd_ind + seite - 1].He_Nachname + '</td></tr>';
            } else {
                HTML_Seite = '<tr id = "anzeige_body"><td>' + runden_info[rd_ind + seite - 1].Name_Team + '</td></tr>';
            }
            HTML_Seite += '<tr><td><table style="width: 800px; text-align: center;" cellspacing="0" cellpadding="1" border="1"><tbody><tr bgcolor="#ddd">';
            if (runden_info[0].Startklasse.substring(0, 3) === "F_B") {
                HTML_Seite += '<td>WR</td><td style="width: 120px;">Schritt</td><td style="width: 120px;">Syn</td><td>Fig Au</td><td>Fig Sch</td><td style="width: 120px;">Fig Bo</td><td style="width: 120px;">Ch</td><td style="width: 120px;">Ch Id</td><td style="width: 120px;">Ch Bo</td><td style="width: 120px;">AF Sy</td><td>AF Pos</td><td style="width: 120px;">AF Bo</td><td style="width: 120px;">Summe</td><td>&nbsp;</td></tr>';
            } else {
                if (new_guidelines) {
                    HTML_Seite += '<td>WR</td><td style="width: 120px;">GS D</td><td style="width: 120px;">GS H</td><td>Basic Dance</td><td>Dance Perf</td><td style="width: 120px;">Bonus</td><td style="width: 120px;">Fig Ausf</td><td style="width: 120px;">Fig Schw</td><td style="width: 120px;">Bonus</td><td style="width: 120px;">Interpr</td><td>Spontan</td><td style="width: 120px;">Bonus</td><td style="width: 120px;">Summe</td><td>&nbsp;</td></tr>';
                } else {
                    if (trunde === 'ER') {
                        HTML_Seite += '<td>Name</td><td>Grundschritt</td><td>Basic Dancing</td><td>Tanzfig</td><td>Interpret</td><td>Spontane Int</td><td>Dance Perf</td><td>Summe</td><td>&nbsp;</td></tr>';
                    } else {
                        HTML_Seite += '<td>Name</td><td>Grundschritt</td><td>Tanzfig</td><td>Interpret</td><td>Dance Perf</td><td>Summe</td><td>&nbsp;</td></tr>';
                    }
                }
            } // wr-teil
            for (wr in wertungsrichter) {
//                if (wertungsrichter[wr].WR_Azubi === false) {
                    if (wertungsrichter[wr].WR_func === "X") {
                        tp_id = runden_info[rd_ind + seite - 1].TP_ID;
                        if (typeof wertungen[wr] !== "undefined") {
                            if (typeof wertungen[wr][tp_id] !== "undefined") {
                                cgi_val = wertungen[wr][tp_id].cgi;
                                if (runden_info[0].Startklasse.substring(0, 3) === "F_B") {
                                    if (wertungen[wr][tp_id].in === true) {
                                        HTML_Seite += '<tr bgcolor="#fff"><td height="35px">' + wertungsrichter[wr].WR_Nachname + '</td>';
                                    } else {
                                        HTML_Seite += '<tr bgcolor="#dbd"><td height="35px">' + wertungsrichter[wr].WR_Nachname + '</td>';
                                    }
                                    var wr_kr = ["ng_ttd", "ng_tth", "ng_bda", "ng_dap", "ng_bdb", "ng_fta", "ng_fts", "ng_ftb", "ng_inf", "ng_ins", "ng_inb"];
                                    for (var k = 0; k < wr_kr.length; k++) {
                                        if (wertungen[wr][tp_id]["w" + wr_kr[k] + seite] === false) {
                                            HTML_Seite += '<td bgcolor="#dbd" fld=' + wr_kr[k] + seite + '>' + parseFloat(cgi_val["w" + wr_kr[k] + seite]) + '</td>';
                                        } else {
                                            HTML_Seite += '<td fld=' + wr_kr[k] + seite + '>' + parseFloat(cgi_val["w" + wr_kr[k] + seite]) + '</td>';
                                        }
                                    }
                                } else {
                                    if (wertungsrichter[wr].WR_Azubi === true) {    //  WR_Azubi gelb
                                        HTML_Seite += '<tr bgcolor="#ff0"><td height="35px">' + wertungsrichter[wr].WR_Nachname + '</td>';
                                    } else { 
                                        if (wertungen[wr][tp_id].in === true) {
                                            HTML_Seite += '<tr bgcolor="#fff"><td height="35px">' + wertungsrichter[wr].WR_Nachname + '</td>';
                                        } else {
                                            HTML_Seite += '<tr bgcolor="#dbd"><td height="35px">' + wertungsrichter[wr].WR_Nachname + '</td>';
                                         }
                                    }
                                    if (new_guidelines) {
                                        var wr_kr = ["ng_ttd", "ng_tth", "ng_bda", "ng_dap", "ng_bdb", "ng_fta", "ng_fts", "ng_ftb", "ng_inf", "ng_ins", "ng_inb"];
                                        for (var k = 0; k < wr_kr.length; k++) {
                                            if (wertungen[wr][tp_id]["w" + wr_kr[k] + seite] === false) {
                                                HTML_Seite += '<td bgcolor="#dbd" fld=' + wr_kr[k] + seite + '>' + parseFloat(cgi_val["w" + wr_kr[k] + seite]) + '</td>';
                                            } else {
                                                HTML_Seite += '<td fld=' + wr_kr[k] + seite + '>' + parseFloat(cgi_val["w" + wr_kr[k] + seite]) + '</td>';
                                            }
                                        }
                                    } else {
                                        HTML_Seite += '<td>' + parseFloat(cgi_val["wgs" + seite]) + '</td>';
                                        if (trunde === 'ER') {
                                            HTML_Seite += '<td>' + parseFloat(cgi_val["wbd" + seite]) + '</td>';
                                        }
                                        HTML_Seite += '<td>' + parseFloat(cgi_val["wtf" + seite]) + '</td>';
                                        HTML_Seite += '<td>' + parseFloat(cgi_val["win" + seite]) + '</td>';
                                        if (trunde === 'ER') {
                                            HTML_Seite += '<td>' + parseFloat(cgi_val["wsi" + seite]) + '</td>';
                                        }
                                        HTML_Seite += '<td>' + parseFloat(cgi_val["wdp" + seite]) + '</td>';
                                    }
                                }
                                if (runden_info[0].Runde.indexOf("r_schnell") > 0) {
                                    HTML_Seite += '<td>' + fix2(wertungen[wr][tp_id].Punkte / 1.1) + '</td>';
                                } else {
                                    HTML_Seite += '<td>' + fix2(wertungen[wr][tp_id].Punkte) + '</td>';
                                }
                                HTML_Seite += '<td><input class="wr_nochmal" value="nochmal werten" type="button" onclick="senden(this.value, ' + wr + ')"></td ></tr>';
                            }
                        }
                    }
                    if (wertungsrichter[wr].WR_func === "Ob") {
                        tp_id = runden_info[rd_ind + seite - 1].TP_ID;
                        if (typeof wertungen[wr] !== "undefined") {
                            if (typeof wertungen[wr][tp_id] !== "undefined") {
                                cgi_val = wertungen[wr][tp_id].cgi;
                                HTML_Seite += '<tr><td height="40px">' + wertungsrichter[wr].WR_Nachname + '</td>';
                                HTML_Seite += '<td' + iif('Akro ', cgi_val["wakrovw" + seite]) + '</td>';
                                HTML_Seite += '<td' + iif('Fig Jun ', cgi_val["wjuniorvw" + seite]) + '</td>';
                                HTML_Seite += '<td' + iif('Kleid ', cgi_val["wkleidungvw" + seite]) + '</td>';
                                HTML_Seite += '<td' + iif('Tanzber ', cgi_val["wtanzbereichvw" + seite]) + '</td>';
                                //if (trunde === 'ER') {
                                HTML_Seite += '<td' + iif('TanzZeit ', cgi_val["wtanzzeitvw" + seite]) + '</td>';
                                HTML_Seite += '<td' + iif('Dis ', cgi_val["waufrufvw" + seite]) + '</td>';
                                //}
                                if (new_guidelines) {
                                    HTML_Seite += '<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>';
                                }
                                HTML_Seite += '<td>' + fix2(wertungen[wr][tp_id].Punkte) * -1 + '</td>';
                                HTML_Seite += '<td><input class="wr_nochmal" value="nochmal werten" type="button" onclick="senden(this.value, ' + wr + ')"></td ></tr>';
                            }
                        }
                    }
//                }
            }
            HTML_Seite += '</tbody></table></td></tr><tr><td>&nbsp;<input name="Obs_check' + seite + '" value="Ok" type="hidden"></td></tr>';
            HTML_Arr[seite] = HTML_Seite;
        }
        HTML_Seite = "";
        for (seite in HTML_Arr) {
            HTML_Seite += HTML_Arr[seite];
        }
        HTML_Seite += make_absenden(true, true) + '</table></center></form>';
    } else {
        HTML_Seite = HTML_Seite + '<tr><td>Runde ist zuende</td></tr>';
    }
    io.sockets.emit('chat', { msg: 'body', WR: wr_id, HTML: HTML_Seite, ausw: 'OB_' });
};

function iif(text, wert) {
    if (wert === "") {
        return '>&nbsp;';
    } else {
        if (wert === "0") {
            return ' style = "background: yellow;">' + text;
        } else {
            return ' style = "background: yellow;">' + text + ' ' + to_zahl(wert);
        }
    }
}

exports.RR_Seite = function (rd_ind, runden_info, akrobatiken, wr_func, wr_name, wr_id, io) {
    var st_kl = runden_info[0].Startklasse;
    var trunde = runden_info[0].RundeArt;
    var pkt;
    var akro;

    HTML_Seite = make_kopf(rd_ind, runden_info, 2, wr_name) + '\r\n';
    HTML_Seite += '<tr id="anzeige_body">' + '\r\n';
    for (var s = 1; s <= runden_info[rd_ind].PpR; s++) {
        //------------------------------------Tanzwertungsrichter
        HTML_Seite += '<td class="main"><table align="center" border="0" cellpadding="0" cellspacing="0">' + '\r\n';
        if (wr_func === "Ft") {
            if (trunde === 'ER') {   //      && runden_info[0].Runde != "Semi"
                HTML_Seite += make_inpRRE('sh' + s, 10, 'Technik Herr - Grundtechnik', st_kl) + '\r\n';
                HTML_Seite += make_inpRRE('th' + s, 10, 'Technik Herr - Haltungs- und Drehtechnik', st_kl) + '\r\n';
                HTML_Seite += make_inpRRE('sd' + s, 10, 'Technik Dame - Grundtechnik', st_kl) + '\r\n';
                HTML_Seite += make_inpRRE('td' + s, 10, 'Technik Dame - Haltungs- und Drehtechnik', st_kl) + '\r\n';
                HTML_Seite += make_inpRRE('ch' + s, 10, 'Tanz - Wertigkeit', st_kl) + '\r\n';
                HTML_Seite += make_inpRRE('tf' + s, 10, 'Tanz - Ausführung', st_kl) + '\r\n';
                HTML_Seite += make_inpRRE('da' + s, 10, 'Tanz - Wirkung', st_kl) + '\r\n';
            } else {
                HTML_Seite += make_inpRRE('sh' + s, 10, 'Technik Herr', st_kl) + '\r\n';
                HTML_Seite += make_inpRRE('sd' + s, 10, 'Technik Dame', st_kl) + '\r\n';
                HTML_Seite += make_inpRRE('ch' + s, 10, 'Tanz', st_kl) + '\r\n';
            }
            HTML_Seite += '<tr><td height="12px"></td></tr>' + make_fehler(s, true, false);
        } else {
            for (var t = 1; t <= 6; t++) {
                if (runden_info[rd_ind + s - 1]["Akro" + t + '_' + trunde] !== null) {
                    pkt = akrobatiken[runden_info[rd_ind + s - 1]["Akro" + t + '_' + trunde]][st_kl].replace(/,/, ".");
                    akro = akrobatiken[runden_info[rd_ind + s - 1]["Akro" + t + '_' + trunde]]['Langtext'];
                    HTML_Seite += make_inpRRE('ak' + s + t, pkt, akro, st_kl);
                } else {
                    HTML_Seite += '<tr><td></td></tr>';
                    HTML_Seite += '<tr><td width="15" height="15"></td></tr>';
                }
            }
        }
        HTML_Seite += '</table></td>' + '\r\n';
    }
    HTML_Seite += '</tr>' + make_absenden(true, false) + '</table></center></form>';

    io.sockets.emit('chat', { msg: 'body', WR: wr_id, HTML: HTML_Seite, ausw: 'RR_' });
};

exports.RR_Observer = function (rd_ind, runden_info, seite, wr_name, wr_id, akrobatiken, anz_obs, io) {
    var st_kl = runden_info[0].Startklasse;
    var trunde;
    var ausw = "RR_";
    var btn_aktiv = true;
    var seiten = 1 + seite;
    var s = 1 + seite;
    var rd_info;
    if (anz_obs === 1 && runden_info[rd_ind].PpR === 2) {
        seiten = 2;
    }
    HTML_Seite = make_kopf(rd_ind + seite, runden_info, seiten - s + 1, wr_name, false, s === 2 && seiten === 2) + '\r\n';
    HTML_Seite += '<tr id="anzeige_body">' + '\r\n';
    //----
    for (s; s <= seiten; s++) {
        HTML_Seite += '<td class="main"><table align="center" border="0" cellpadding="0" cellspacing="0">' + '\r\n';
        var rde = (runden_info[0].Runde.substr(0, 4));
        if (st_kl === "RR_S1" || st_kl === "RR_S2" || rde === "MK_1" || rde === "MK_2" || rde === "MK_3" || rde === "MK_4") {
            switch (rde) {
                case "MK_1":
                case "MK_2":
                    HTML_Seite += '<tr><td height="10"></td></tr>' + '\r\n';
                    HTML_Seite += make_inpMKText('mk_td' + s, 0, "Dame") + '\r\n';
                    HTML_Seite += '<tr><td height="20"></td></tr>' + '\r\n';
                    HTML_Seite += make_inpMKText('mk_th' + s, 0, "Herr") + '\r\n';
                    HTML_Seite += '<tr><td height="150"></td></tr>' + '\r\n';
                    ausw = "MK_T";
                    break;
                case "MK_3":
                case "MK_4":
                    HTML_Seite += '<tr><td height="50"></td></tr>' + '\r\n';
                    HTML_Seite += make_inpMK('mk_td' + s, 10, 'Dame', st_kl) + '\r\n';
                    HTML_Seite += '<tr><td height="20"></td></tr>' + '\r\n';
                    HTML_Seite += make_inpMK('mk_th' + s, 10, 'Herr', st_kl) + '\r\n';
                    HTML_Seite += '<tr><td height="50"></td></tr>' + '\r\n';
                    ausw = "MK_A";
                    break;
                case "MK_5":
                    HTML_Seite += '<tr><td height="100"></td></tr>';
                    HTML_Seite += make_inpMK('mk_th' + s, 10, 'Figuren', st_kl) + '\r\n';
                    HTML_Seite += '<tr><td height="100"></td></tr>';
                    ausw = "MK_A";
                    break;
            }
            btn_aktiv = false;
        } else {
            var trunde = runden_info[0].RundeArt;
            if (runden_info[rd_ind].Runde.indexOf("_Fu") > 0 || st_kl === "RR_S") {
                HTML_Seite += '<tr><td width="430" height="200"></td></tr>';
            } else {
                rd_info = rd_ind + s - 1;        //parseInt(s / (anz_obs - 1)) - 1;
                for (var t = 1; t <= 8; t++) {
                    pkt = runden_info[rd_info]['Wert' + t + '_' + trunde];
                    if (pkt !== null) {
                        HTML_Seite += '<tr><td class="akro" width="430">' + akrobatiken[runden_info[rd_info]["Akro" + t + '_' + trunde]]['Langtext'] + '</td></tr>' + make_fehler('ak' + s + t, false, true);
                    } else {
                        HTML_Seite += '<tr><td></td></tr>';
                        HTML_Seite += '<tr><td width="15" height="15"></td></tr>';
                    }
                }
            }
            HTML_Seite += '<tr><td class="akro" width="400px">Fu&szlig;technik</td></tr>';
            HTML_Seite += make_fehler(s, true, false);
            HTML_Seite += make_A20(s) + '\r\n';
        }
        HTML_Seite += '</table></td>' + '\r\n';
    }

    HTML_Seite += '</tr>' + make_absenden(true, btn_aktiv);
    HTML_Seite += '</table></center></form>';
    HTML_Seite += '<script> const ausw="' + runden_info[0].Startklasse.substr(0, 2) + '" </script>';

    io.sockets.emit('chat', { msg: 'body', WR: wr_id, HTML: HTML_Seite, ausw: ausw });
};

exports.RR_ObsCheck = function (rd_ind, wertungsrichter, wertungen, runden_info, runde, seite, wr_name, wr_id, akrobatiken, anz_obs, io) {
    var PunkteOb = new Array();
    var PunkteFt = new Array();
    var PunkteAk = new Array();
    var HTML_name;
    var HTML_Fehler;
    var HTML_array = new Object();
    var cgi_val = new Object();
    var trunde = runden_info[0].RundeArt;
    var wr;
    var wrn;
    var c;
    var but_send = false;
    var seiten = runden_info[rd_ind].PpR;
    var s = 1;
    if (anz_obs === 2) {
        seiten = seite + 1;
        s = seite + 1;
    }
    var tp_id;
    HTML_Seite = make_kopf(rd_ind + seite, runden_info, 0, wr_name) + '\r\n';
    HTML_Seite += '<tr><td class="akro"><input name="Obs_check1" value="Ok" type="hidden"><input type="hidden" name=rh1 value="' + runden_info[rd_ind].Rundennummer + '"><input name="rt_ID" value="' + runden_info[rd_ind].RT_ID + '" type="hidden"><table colspan="10" width="100%" rules="rows">' + '\r\n';

    for (s; s <= seiten; s++) {
        PunkteOb = new Array();
        PunkteFt = new Array();
        PunkteAk = new Array();
        HTML_array = new Object();
        tp_id = runden_info[rd_ind + seite].TP_ID;
        for (wr in wertungen) {
            if (typeof wertungen[wr][tp_id] !== "undefined") {
                if (wertungsrichter[wr].WR_Azubi === false) {
                    wrn = wertungsrichter[wr].WR_Nachname.substring(0, 1) + wertungsrichter[wr].WR_Vorname.substring(0, 2) + wertungsrichter[wr].WR_ID;
                    switch (wertungsrichter[wr].WR_func) {
                        case 'Ob':
                            PunkteOb = { "WR_ID": wr, "WR_func": wertungsrichter[wr].WR_func, "rd_ind": rd_ind + seite, "paar": wertungen[wr], "seite": s, "WR_name": wrn, "cgi": wertungen[wr][tp_id].cgi };
                            break;
                        case 'Ft':
                            PunkteFt[wr] = { "WR_ID": wr, "WR_func": wertungsrichter[wr].WR_func, "rd_ind": rd_ind + seite, "paar": wertungen[wr], "seite": s, "WR_name": wrn, "cgi": wertungen[wr][tp_id].cgi };
                            break;
                        case 'Ak':
                            PunkteAk[wr] = { "WR_ID": wr, "WR_func": wertungsrichter[wr].WR_func, "rd_ind": rd_ind + seite, "paar": wertungen[wr], "seite": s, "WR_name": wrn, "cgi": wertungen[wr][tp_id].cgi };
                            break;
                    }
                }
            }
        }
        var korr = false;
        // Fußtechnik WR erstellen --------------------------
        HTML_name = '<tr id="anzeige_inhalt' + s + '"><td><b> Startnummer : ' + runden_info[rd_ind + s - 1].Startnr + '</b><input type="hidden" name=TP_ID' + s + ' value="' + runden_info[rd_ind + s - 1].TP_ID + '"></td>';
        HTML_fehler = '<tr><td>Fu&szlig;technik</td>';
        for (i in PunkteFt) {
            cgi_val = PunkteFt[i].cgi;
            HTML_name += '<td>' + PunkteFt[i].WR_name.substring(0, 3) + '</td>';
            if (cgi_val['wfl' + s] > 0) {
                HTML_fehler += '<td>' + get_grobfehler(cgi_val['tfl' + s], 'fl' + s) + '</td>';
                korr = true;
            } else {
                HTML_fehler += '<td>-</td>';
            }
        }
        // Fußtechnik Ob erstellen --------------------------
        cgi_val = PunkteOb.cgi;
        HTML_name += '<td>' + PunkteOb.WR_name.substring(0, 3) + '</td>';
        if (cgi_val['wfl' + s] > 0) {
            HTML_fehler += '<td>' + get_grobfehler(cgi_val['tfl' + s], 'fl' + s) + '</td>';
            korr = true;
        } else {
            HTML_fehler += '<td>-</td>';
        }
        // Fußtechnik Korrektur erstellen --------------------------
        HTML_name += '<td>Korrektur</td></tr>';
        if (korr === true) {
            HTML_fehler += '<td class="wr_status div"><table><tr id="fl' + s + '">';
            HTML_fehler += '<td class="obsbuttons">&nbsp;&nbsp;-&nbsp;&nbsp;</td>';
            HTML_fehler += '<td><input id="tfl' + s + '" name="tfl' + s + '" class="mistakes_inputs" autocomplete="off" type="text"><input value="0" type="hidden" name="wfl' + s + '" id="wfl' + s + '"></td>';
            HTML_fehler += '</tr></table>';
        } else {
            HTML_fehler += '<td>-</td>';
        }
        HTML_Seite += HTML_name + HTML_fehler;

        //  Akrobatiken sammeln
        for (var i = 1; i < 9; i++) {
            if (runden_info[rd_ind + s -1]["Akro" + i + '_' + trunde] !== null) {
                HTML_array[i] = new Object();
                HTML_array[i]["Akro"] = runden_info[rd_ind + s - 1]["Akro" + i + '_' + trunde];
            }
        }
        wr = new Array();
        for (i in PunkteAk) {
            cgi_val = PunkteAk[i].cgi;
            for (c in HTML_array) {
                if (cgi_val['tflak' + s + c] === "") {
                    HTML_array[c][PunkteAk[i].WR_name] = "-";
                } else {
                    HTML_array[c][PunkteAk[i].WR_name] = cgi_val['tflak' + s + c];
                    HTML_array[c]["korr"] = true;
                    korr = true;
                }
            }
        }
        cgi_val = PunkteOb.cgi;
        for (c in HTML_array) {
            if (cgi_val['tflak' + s + c] === "" || typeof cgi_val['tflak' + s + c] === "undefined") {
                HTML_array[c][PunkteOb.WR_name] = "-";
            } else {
                HTML_array[c][PunkteOb.WR_name] = cgi_val['tflak' + s + c];
                HTML_array[c]["korr"] = true;
                korr = true;
            }
        }

        if (PunkteAk.length > 0) {
            HTML_Seite += '<tr><td colspan="10" height="20"></td></tr><tr><td clospan="10">Akrobatiken</td></tr>';
            HTML_name = '<tr><td></td>';
            wr = new Array();
            for (c in HTML_array[1]) {
                if (c !== "Akro" && c !== "korr") {
                    HTML_name += '<td>' + c.substring(0,3) + '</td>';
                    wr.push(c);
                }
            }
            HTML_name += '<td>Korrektur</td></tr>' + '\r\n';

            HTML_fehler = "";
            for (i in HTML_array) {
                HTML_fehler += '<tr id="flak' + s + i + '"><td>' + akrobatiken[HTML_array[i]["Akro"]].Langtext + '</td>';
                // Akro WR Ob erstellen --------------------------
                for (c in wr) {
                    HTML_fehler += '<td>' + get_grobfehler(HTML_array[i][wr[c]], 'flak' + s + i) + '</td>';
                }
                if (HTML_array[i]["korr"] === true) {
                    // Akro back erstellen --------------------------
                    HTML_fehler += '<td><table><tr id="flak' + s + i + '">';
                    HTML_fehler += '<td class="obsbuttons">&nbsp;&nbsp;-&nbsp;&nbsp;</td>';
                    HTML_fehler += '<td><input id="tflak' + s + i + '" name="tflak' + s + i + '" class="mistakes_inputs" autocomplete="off" type="text"><input value="0" type="hidden" name="wflak' + s + i + '" id="wflak' + s + i + '"></td>';
                    HTML_fehler += '</tr></table></td>';
                } else {
                    HTML_fehler += '<td>-</td>';
                }
                HTML_fehler += '</tr>' + '\r\n';
            }
            HTML_Seite += HTML_name + HTML_fehler;
        } else {
            HTML_Seite += '<tr><td width="450" height="120"></td></tr>';
        }
        // A20 Z20 Anzeige ----------------------------------------------
        HTML_Seite += '<tr>';
        if (cgi_val['tfl' + s + 'a20'] !== "" && typeof cgi_val['tfl' + s + 'a20'] !== "undefined") {
            HTML_Seite += '<td colspan="' + wr.length + '">A20 Z20</td><td><table><tr id="fl' + s + 'a20">';
            if (cgi_val['tfl' + s + 'a20'].indexOf("A20") > -1) {
                HTML_Seite += '<td class="obsbuttons">A20</td >';
            }
            if (cgi_val['tfl' + s + 'a20'].indexOf("Z20") > -1) {
                HTML_Seite += '<td class="obsbuttons">Z20</td >';
            }
            HTML_Seite += '</tr ></table ></td >';
            // A20 back ----------------------------------------------
            HTML_Seite += '<td><table><tr id="fl' + s + 'a20"><td class="obsbuttons">&nbsp;&nbsp;-&nbsp;&nbsp;</td>';
            HTML_Seite += '<td><input id="tfl' + s + 'a20" name="tfl' + s + 'a20" class="mistakes_inputs" autocomplete="off" type="text" value=""><input value="0" type="hidden" name="wfl' + s + 'a20" id="wfl' + s + 'a20"></td>';
            HTML_Seite += '</tr></table></td>';
            korr = true;
        }
        HTML_Seite += '</tr>';

        HTML_Seite += '<tr><td colspan="10" bgcolor="#000"><input name="korr' + s + '" value="' + korr + '" type="hidden"></td></tr>';
        if (korr === true) {
            but_send = true;
        }
    }
    HTML_Seite += '</table ></td></tr>';
    HTML_Seite += make_absenden(true, !but_send) + '</table></center></form>';

    io.sockets.emit('chat', { msg: 'body', WR: wr_id, HTML: HTML_Seite, ausw: 'RR_' });
};

exports.RR_Form_Seite = function (rd_ind, runden_info, akrobatiken, wr_func, wr_name, wr_id, io) {
    var st_kl = runden_info[0].Startklasse;
    var trunde = runden_info[0].RundeArt;
    var akro;

    var HTML_Seite = make_kopf(rd_ind, runden_info, 2, wr_name) + '\r\n';
    HTML_Seite += '<tr id="anzeige_body" style="min-height: 400px;">' + '\r\n';
    HTML_Seite += '<td class="main"><table align="center" border="0" cellpadding="0" cellspacing="0">' + '\r\n';
    switch (wr_func) {
        case "X":
            HTML_Seite += make_inpRRE('tk1', 10, 'Tanztechnik - Grundschritt (Rhythmus und Fußtechnik) / Basic Dancing', st_kl) + '\r\n';
            HTML_Seite += make_inpRRE('ch1', 10, 'Tanzfiguren - komplexe Figuren, Highlightfiguren', st_kl) + '\r\n';
            HTML_Seite += make_inpRRE('tf1', 10, 'Choreografie / Dance Performance (Aufbau, Musikinterpretation, Präsentation, Ausstrahlung', st_kl) + '\r\n';
            HTML_Seite += make_inpRRE('ab1', 10, 'AF - Synchronität und Harmonie', st_kl) + '\r\n';
            HTML_Seite += make_inpRRE('aw1', 10, 'AF - Bilder und Bildwechsel (Schwierigkeit, Ausführung)', st_kl) + '\r\n';
            HTML_Seite += make_inpRRE('af1', 10, 'AF - Formationsfiguren und Effekte', st_kl) + '\r\n';
            HTML_Seite += '<tr><td height="12px"></td></tr>';
            break;
        case "Ft":
            HTML_Seite += make_inpRRE('tk1', 10, 'Technik - Grund-, Haltungs- und Drehtechnik', st_kl) + '\r\n';
            HTML_Seite += make_inpRRE('ch1', 10, 'Tanz - Wert inkl. Formationsfiguren und Abstimmung zur Musik', st_kl) + '\r\n';
            HTML_Seite += make_inpRRE('tf1', 10, 'Tanz - Ausführung', st_kl) + '\r\n';
            HTML_Seite += make_inpRRE('ab1', 10, 'AF - Wert der Bilder, Bildwechsel und Effekte', st_kl) + '\r\n';
            HTML_Seite += make_inpRRE('aw1', 10, 'AF - Ausführung', st_kl) + '\r\n';
            HTML_Seite += make_inpRRE('af1', 10, 'Gesamtwirkung', st_kl) + '\r\n';
            HTML_Seite += '<tr><td height="12px"></td></tr>' + make_fehler(1, true, false);
            break;
        case "Ak":
            for (var t = 1; t <= 8; t++) {
                pkt = runden_info[rd_ind]['Wert' + t + '_' + trunde];
                if (pkt !== null) {
                    akro = akrobatiken[runden_info[rd_ind]["Akro" + t + '_' + trunde]]['Langtext'];
                    pkt = akrobatiken[runden_info[rd_ind]["Akro" + t + '_' + trunde]][st_kl].replace(",", ".");
                    HTML_Seite += make_inpRRE('ak1' + t, pkt, akro, st_kl);
                }
            }
            HTML_Seite += '<tr><td></td></tr>';
            break;
        default:
    }
    HTML_Seite += '</table></td></tr>' + '\r\n';
    HTML_Seite += make_absenden(true, false) + '</table></center></form>';
    HTML_Seite += '<script> const ausw="' + runden_info[0].Startklasse.substr(0, 2) + '" </script>';

    io.sockets.emit('chat', { msg: 'body', WR: wr_id, HTML: HTML_Seite, ausw: 'RR_' });
};

exports.reisekosten = function (runden_info, wr_id, io, connection) {
    connection
        .query('SELECT * FROM Wert_Richter WHERE WR_ID =' + wr_id + ' ;')
        .on('done', function (data) {
            var HTML_Seite = make_kopf(0, runden_info, 0, data[0].WR_Nachname.substr(0, 1) + data[0].WR_Vorname, false) + '\r\n';
            HTML_Seite += '<tr><td style="font-family: Arial; font-size: 16px; width: 50%; "><table align="center" border="0"><tbody>';
            // <!-- personal Daten -->
            HTML_Seite += make_reise('<b>Wertungsrichter</b>', '&nbsp;');
            HTML_Seite += make_reise('Name:', data[0].WR_Vorname + ' ' + data[0].WR_Nachname);
            HTML_Seite += make_reise('&nbsp;', '&nbsp;');
            HTML_Seite += make_reise('Reise von/nach:', ' nach Zwönitz');
            HTML_Seite += make_reise('Grund der Reise:', '<input class="reise_i" style="width:300px;" id="r_grund" value="' + runden_info[0].Turnier_Name + '">');
            HTML_Seite += '<tr> <td height="10"></tr>';
            HTML_Seite += make_reise('Beginn der Reise:', '<input class="reise_i" id="r_beginn" value="01.10.2022 09:30">');
            HTML_Seite += make_reise('Ende der Reise:', '<input class="reise_i" id="r_ende" value="02.10.2022 01:30">');
            HTML_Seite += '<tr> <td height="10"></tr>';
            HTML_Seite += make_reise('Beginn des Dienstgeschäfts:', '<input class="reise_i" id="r_dbeginn" value="01.10.2022 09:30">');
            HTML_Seite += make_reise('Ende des Dienstgeschäfts:', '<input class="reise_i" id="r_dende" value="02.10.2022 01:30">');
            HTML_Seite += '<tr> <td height="10"></tr>';
            //	<!-- Fahrkosten  -->
            HTML_Seite += make_reise('<b>Fahr- und Nebenkosten</b>', '&nbsp;');
            HTML_Seite += make_reise('Auto - einfache Entfernung:', '<input class="reise_i" onchange="reise_summe()" style="width:70px;" id="r_km" value="330"> km');
            HTML_Seite += make_reise('andere Anreise', '<input class="reise_i" style="width:200px;" type="text" id="r_bahn" value="" placeholder="z.B. Bahn / Flug">');
            HTML_Seite += make_reise('alle Ausgaben', '<input onchange="reise_summe()" style="width:100px;" class="reise_i" id="r_kbahn">&nbsp;€');
            HTML_Seite += '<tr> <td height="10"></tr>';
            HTML_Seite += '</tbody></table></td>';
            HTML_Seite += '<td style="font-family: Arial; font-size: 16px;"><table align="center" border="0"><tbody>';
            //	<!-- Verpflegungskosten  -->
            HTML_Seite += make_reise('<b>Verpflegungspauschale</b>', '&nbsp;');
            HTML_Seite += make_reise('Abwesenheit von mehr als 8 Stunden:', '<input onchange="reise_summe()" style="width:50px;" class="reise_i" id="r_pausch14">&nbsp;Tage');
            HTML_Seite += make_reise('mehrtägiger Auswärtstätigkeit:', '<input onchange="reise_summe()" style="width:50px;" class="reise_i" id="r_pausch28">&nbsp;Tage');
            HTML_Seite += '<tr> <td height="10"></tr>';
            HTML_Seite += make_reise('abzgl. € 5,60 je Frühstück:', '<input onchange="reise_summe()" style="width:50px;" class="reise_i" id="r_frueh">&nbsp;Tage');
            HTML_Seite += make_reise('abzgl. € 11,20 je Mittagessen:', '<input onchange="reise_summe()" style="width:50px;" class="reise_i" id="r_essen">&nbsp;Tage');
            HTML_Seite += make_reise('abzgl. € 11,20 je Abendessen:', '<input onchange="reise_summe()" style="width:50px;" class="reise_i" id="r_abend">&nbsp;Tage');
            HTML_Seite += '<tr> <td height="10"></tr>';
            HTML_Seite += make_reise('Übernachtungskosten:', '<input onchange="reise_summe()" style="width:100px;" class="reise_i" id="r_uekosten">&nbsp;€');
            HTML_Seite += '<tr> <td height="10"></tr>';
            //	<!-- Honorar  -->
            HTML_Seite += make_reise('<b>Honorar</b>', '<input style="width:250px;" class="reise_i" id="r_hono" value="" placeholder="z.B. Moderator / DJ">');
            HTML_Seite += make_reise('Honorarkosten:', '<input onchange="reise_summe()" style="width:100px;" class="reise_i" id="r_khono">&nbsp;€');
            HTML_Seite += '<tr> <td height="10"></tr>';
            //  <!-- Summe  -->	
            HTML_Seite += make_reise('<b>Gesammtsumme</b>', '<input style="width:100px;" class="reise_i" id="r_summe" readonly>&nbsp;€');
            HTML_Seite += '</tbody></table></td></tr>';
//            HTML_Seite += '</tr>' + make_absenden(true, true, false) + '</table></center></form>';
            HTML_Seite += '<tr id = "anzeige_absenden"><td class="button_b" colspan="4">';
            HTML_Seite += '<input id= "absend" name= "absend" class="button_2" value= "Absenden" onclick = "reise_send()" type = "button" > ';
            HTML_Seite += '<input id="wtim" name="wtim" type="hidden"></td></tr>';

            io.sockets.emit('chat', { msg: 'body', WR: wr_id, HTML: HTML_Seite, ausw: 'BW_' });
        });

    function make_reise(w1, w2) {
        var HTML_Seite = '<tr><td class="reise_r">' + w1 + '</td >';
        HTML_Seite += '<td><td class="reise_l">' + w2 + '</td></tr >';
        return HTML_Seite;
    }
};

function get_grobfehler(fl, fl_id) {
    var gr = fl.trim().split(' ');
    var HTML = '<table><tr id="' + fl_id + '">';
    if (gr[0] === "-") {
            HTML += '<td>-</td>';
    } else {
        for (var i in gr) {
            HTML += '<td class="obsbuttons">' + gr[i] + '</td>';
        }
    }
    HTML += '</tr></table>';
    return HTML;
}

function make_inpRRE(fName, max, aName, st_kl) {
    var ak;
    var t;
    if (fName.indexOf("ak") === 0) {
        ak = '<tr><td class="akro" colspan = "21">' + aName + '</td></tr><tr>' + make_fehler(fName, false, true);
    } else {
        ak = '<tr><td class="akro" colspan = "21">' + aName + '</td></tr>';
    }
    ak += '<tr id="' + fName + '" class="kriterium">';
    for (t = 0; t < 21; t++) {
        ak += '<td class="btn_wert">' + t * 5 + '</td>';
    }
    ak += '<td><input id="w' + fName + '" value="" type="hidden" name="w' + fName + '" max="' + max + '"></td></tr>';
    return ak;
}

function make_inpBW(fName, max, aName, st_kl) {
    var inp;
    inp = '<div class="schrift">' + aName + '</div>';
    inp += '<div class="kriterium" id="' + fName + '">';
    for (var t = 0; t < max + 1; t++) {
        inp += '<div class="btn_leer">' + t + '</div>' + '\r\n';
//  IIf(wert % 2, "-", wert / 2)
    }
    inp += '</div>';
    inp += '<input name="w' + fName + '" id="w' + fName + '" type="hidden">';

    return inp;
}

function make_inpNG_BW(fName, max, aName, st_kl, vorbelegung) {
    var inp;
    inp = '<table cellspacing="0" align="center"><tr id="' + fName + '" class="kriterium_NG">';
    for (var t = 0; t < max * 2; t++) {
        inp += '<td class="btn_NG_leer">';
        if (t % 2) {
            inp += '-' + '</td>';
        } else {
            inp += t / 2 + '</td>';
        }
    }
    inp += '<td align="center" class="btn_NG_leer">' + t / 2 + '<input id="w' + fName + '" type="hidden" name="w' + fName + '" value="' + vorbelegung + '"></td>';
    inp += '</tr></table>';

    return inp;
}

function make_inpMKText(fName, max, aName, st_kl) {
    var inp;
    inp = '<tr><td colspan="20">' + aName + '</td></tr>';
    inp += '<tr class="mk_inp" id="' + fName + '"><td><input class="mk_fld" id="w' + fName + '" name="w' + fName + '"  autocomplete="off" onkeyup="wr_onclick(event)"></td></tr>';

    return inp;
}

function make_inpMK(fName, max, aName, ganz, pre, paar) {
    var inp;
    var b_class;
    if (pre === undefined) {
        //        pre = -1;
    }
    var g_zahl = parseInt(pre);
    var k_zahl = (parseFloat(pre) - g_zahl).toFixed(1);
    inp = '<tr class="bs_head"><td><table align="center"><tr><td colspan="20">' + aName + '</td></tr>';
    inp += '<tr class="bs_krit ' + paar + '" id="' + fName + '" max="' + max + '">';
    if (max === 10) {
        for (var t = 0; t < max + 1; t++) {
            if (g_zahl === t) {
                b_class = "mk_bsel";
            } else {
                b_class = "mk_bwert";
            }
            inp += '<td class="' + b_class + '">' + t + '</td>';
        }
    } else {
        for (var t = 0; t < 8; t++) {
            if (g_zahl === t) {
                b_class = "mk_bsel";
            } else {
                b_class = "mk_bwert";
            }
            if (t < max + 1) {
                inp += '<td class="' + b_class + '">' + t + '</td>';
            } else {
                inp += '<td class = "mk_binvi"> </td>';
            }
        }
        inp += '<td class="mk_bwert">+</td>';
        for (t = 1; t < 5; t++) {
            var k_wert = (t * 0.2).toFixed(1);
            if (k_zahl === k_wert) {
                b_class = "mk_bsel";
            } else {
                b_class = "mk_bwert";
            }
            inp += '<td class="' + b_class + '">' + k_wert + '</td>';
        }
    }
    inp += '<td width="30px"> </td>';
    if (pre === undefined) {
        pre = '';
    }
    inp += '<td width="30px"><input class="mk_berg" name="w' + fName + '" id="w' + fName + '" value="' + pre + '" disabled="true"></td></tr>';
    inp += '</tr></table></td>';
    return inp;
}

function kriterium_text(text1, text2) {
    var inp = '<tr>';
    if (typeof text2 !== "undefined") {
        inp += '<td style="padding-top: 15px;padding-bottom: 5px;">' + text1 + '</td>';
        inp += '<td style="padding-top: 15px;padding-bottom: 5px; padding-left: 15px;">' + text2 + '</td></tr>';
    } else {
        inp += '<td style="padding-top: 15px;padding-bottom: 5px;" colspan="3">' + text1 + '</td></tr>';
    }

    return inp;
}

function make_kopf(rd_ind, runden_info, seiten, wr_name, tausch, switch_TP) {
    var HTML_kopf;
    var HTML_nav;
    var sei1 = rd_ind;
    var sei2 = rd_ind + 1;
    if (tausch === true && runden_info[rd_ind].PpR === 2) {
        sei1 = rd_ind + 1;
        sei2 = rd_ind;
    }
    HTML_nav = '<td align="center" width="40"><nav><ul><li onclick="getPaging(event)">&#x2630;<ul id="nav_innen">';
    HTML_nav += '<li id = "li_zeit"> Zeitplan</li>';
    if (rei) {
        HTML_nav += '<li id = "li_reise"> Reisekosten</li>';
        HTML_nav += '<li id = "li_back"> zurück</li>';
    }
    HTML_nav += '</ul ></li ></ul ></nav ></td > ' + '\r\n';

    HTML_kopf = '<form name="Formular" action=/judge method=post onsubmit="return chkFormular()"><center >';
    HTML_kopf += '<table frame="void" border="1" bordercolor="#888888" cellpadding="0" cellspacing="0">';
    HTML_kopf += '<tr id="anzeige_kopf"><td class="rd_o" colspan="4"><table width="100%" cellspacing="0" cellpadding="0" border="0"><tr>';
    if (seiten === 0) {
        HTML_kopf += '<td class="kopf_1" width="50"></td>';
        HTML_kopf += '<td class="runden" width="400"></td>' + '\r\n';
        HTML_kopf += HTML_nav + '\r\n';
        HTML_kopf += '<td style="padding-left: 10px;" onclick="return p_logout()" style="padding-left: 10px;" width="300">Logout ' + wr_name + '</td>' + '\r\n';
        HTML_kopf += '<td class="kopf_1" width="50"></td>' + '\r\n';
    } else {
        if (switch_TP === true) {
            HTML_kopf += '<td class="kopf_1" width="50">' + runden_info[sei1].Startnr + '<input type="hidden" name=TP_ID' + (sei1 - rd_ind + 2) + ' value="' + runden_info[sei1].TP_ID + '">';
        } else {
            HTML_kopf += '<td class="kopf_1" width="50">' + runden_info[sei1].Startnr + '<input type="hidden" name=TP_ID' + (sei1 - rd_ind + 1) + ' value="' + runden_info[sei1].TP_ID + '">';
        }
        HTML_kopf += '<input type="hidden" name=rh1 value="' + runden_info[sei1].Rundennummer + '"><input name="rt_ID" value="' + runden_info[sei1].RT_ID + '" type="hidden"></td>' + '\r\n';
        HTML_kopf += '<td class="runden" width="400">' + runden_info[sei1].Tanzrunde_Text + '</td>' + '\r\n';
        HTML_kopf += HTML_nav + '\r\n';
        HTML_kopf += '<td onclick="return p_logout()" style="padding-left: 10px;" width="350">Logout ' + wr_name + '</td>' + '\r\n';
        if (runden_info[sei1].PpR === 2 && seiten === 2) {
            HTML_kopf += '<td class="kopf_1" width="50">' + runden_info[sei2].Startnr + '<input type="hidden" name=TP_ID' + (sei2 - rd_ind + 1) + ' value="' + runden_info[sei2].TP_ID + '"></td>' + '\r\n';
        } else {
            HTML_kopf += '<td class="kopf_1" width="50"></td>' + '\r\n';
        }
    }
    HTML_kopf += '</tr></table></td></tr>' + '\r\n';
    return HTML_kopf;
}

function make_fehler(seite, takt, ak) {
    var ftfl;
    ftfl = '<tr><td colspan="21"><div class="mistakes" id="mistakes' + seite + '">';
    if (takt === true) {
        ftfl += '<div><div class="btn-warning">T2</div><div class="btn-warning">T10</div><div class="btn-warning">T20</div></div>';
    }
    ftfl += '<div><div class="btn-warning">U2</div><div class="btn-warning">U10</div><div class="btn-warning">U20</div></div>';
    ftfl += '<div><div class="btn-warning">S20</div></div>';
    ftfl += '<div><div class="btn-warning">V5</div></div>' + '\r\n';
    if (ak === true) {
        ftfl += '<div><div class="btn-attention">P0</div></div>' + '\r\n';
    }
    ftfl += '<div><div class="mistakes-list" id="mistakes-list' + seite + '"></div></div>' + '\r\n';
    ftfl += '<input name="tfl' + seite + '" id="tfl' + seite + '" type="hidden" value="" autocomplete="off">';
    ftfl += '<input name="wfl' + seite + '" id="wfl' + seite + '" type="hidden" value="0" autocomplete="off"></div></td></tr>';

    return ftfl;
}

function make_A20(seite) {
    var a20fl = '<tr><td colspan="21" style="display: inline-flex;">';
    a20fl += '<div class="mistakes" id = "mistakes' + seite + 'a20";"> ';
    a20fl += '<div><div class="btn-attention">A20</div></div>';
    a20fl += '<div><div class="btn-attention">Z20</div></div>';
    a20fl += '<div><div class="mistakes-list" id="mistakes-list' + seite + 'a20"></div></div>' + '\r\n';
    a20fl += '<input name="tfl' + seite + 'a20" id="tfl' + seite + 'a20" type="hidden" value="" autocomplete="off">';
    a20fl += '<input name="wfl' + seite + 'a20" id="wfl' + seite + 'a20" type="hidden" value="0" autocomplete="off"></div>';
    a20fl += '</td></tr>';

    return a20fl;
}

function make_absenden(button, aktiv, tausch) {
    var sei1 = 1;
    var sei2 = 2;
    if (tausch === true) {
        sei1 = 2;
        sei2 = 1;
    }
    var abs = '<tr id="anzeige_absenden"><td class="unten" colspan="4"><table align="center" width="100%"><tbody><tr>';
    abs += '<td class="wr_info" id="WR-Info' + sei1 + '" width="25%"></td>' + '\r\n';
    abs += '<td class="button_b" colspan="2" width="50%">';
    if (button === true) {
        if (aktiv === true) {
            abs += '<input id= "absend" name= "absend" class="button_2" value= "Absenden"';
        } else {
            abs += '<input id= "absend" name= "absend" class="button_1" value= "Absenden" disabled=""';
        }
        abs += ' onclick="f_send()" type="button">';
    }
    abs += '</td><td class="wr_info" id="WR-Info' + sei2 + '" width="25%"></td>' + '\r\n';
    abs += '<td><input id="wtim" name="wtim" type="hidden"></td>' + '\r\n';
    abs += '</tr></tbody></table></td></tr>';

    return abs;
}

function fix2(wert) {
    var pu = (Math.round(wert * 100) / 100).toString();
    return pu.replace(".", ",");
}

function to_zahl(wert) {
    if (isNaN(wert) || wert == "") {
        return 0;
    } else {
        return parseFloat(wert);
    }
}

function make_HTMLhead(wr_id, runden_info, html_title) {
    var HTML_Seite;

    HTML_Seite = '<!DOCTYPE html>';
    HTML_Seite += '<head><title>' + html_title + '</title><meta http-equiv="expires" content="0">';

    HTML_Seite += '<link rel="stylesheet" href="EWS3.css">';

    HTML_Seite += '<script src="socket.io/socket.io.js"></script>';
    HTML_Seite += '<script src="EWS3.js" ></script>';
    HTML_Seite += '<script> const WR_ID=' + wr_id + '; </script>' + '\r\n';
    HTML_Seite += '</head><body>';

    return HTML_Seite;
}

function make_bs_inp(fName, max, aName, st_kl) {
    var inp;
    inp  = '<tr><td class="bs_schmal"></td></tr>';
    inp += '<tr class="bs_head"><td colspan="20">' + aName + '</td></tr>';
    inp += '<tr class="bs_krit" id="' + fName + '" max="' + max + '">';
    for (var t = 0; t < max * 2 + 1; t++) {
        if (t % 2) {
            if (st_kl === "Ob") {
                inp += '<td class="bs_wert" style="visibility: hidden;">' + ' ' + '</td>';
            } else {
                inp += '<td class="bs_wert">' + '-' + '</td>';
            }
        } else {
            inp += '<td class="bs_wert">' + t / 2 + '</td>';
        }
    }
    inp += '<input name="w' + fName + '" id="w' + fName + '" value="" type="hidden"></tr>';

    return inp;
}

function make_bs_feh(seite, mit_fehler) {
    var bsfl;
    var t;
    if (mit_fehler) {
        bsfl = '<tr><td class="bs_schmal"></td></tr><tr><td colspan="14"><table>';
        bsfl += '<tr class="bs_head"><td colspan="6"><b>Abz&#252;ge</b> kleiner Fehler</td><td colspan="5">gro&#223;er Fehler</td></tr>';
        bsfl += '<tr id="fe' + seite + '" class="bs_mist_list" couple="' + seite + '">';
        for (t = 0; t < 5; t++) {
            bsfl += '<td class="bs_mist" max="2">&nbsp;</td>';
        }
        bsfl += '<td style="visibility: hidden; width:18px;">&nbsp;</td>';
        for (t = 0; t < 5; t++) {
            bsfl += '<td class="bs_mist" max="5">&nbsp;</td>';
        }
        bsfl += '<td class="bs_mist" style="visibility: hidden;">&nbsp;</td>';
        bsfl += '<input name="wfe' + seite + '" id="wfe' + seite + '" type="hidden" value="0">';
        bsfl += '<input name="tfe' + seite + '" id="tfe' + seite + '" type="hidden"></tr></table>';
    } else {
        bsfl = '<tr><td class="bs_schmal"></td></tr><tr><td  style="height:60px;" colspan="14">';
        bsfl += '<input name="wfe' + seite + '" id="wfe' + seite + '" type="hidden" value="0">';
        bsfl += '<input name="tfe' + seite + '" id="tfe' + seite + '" type="hidden"></td>';
    }
    bsfl += '</td><td colspan="3"><b>Punkte<br>gesamt</b></td><td colspan="4" id="punktefe' + seite + '" class="bs_points">0</td></tr>';
    bsfl += '<tr><td class="bs_schmal"></td></tr>';

    return bsfl;
}