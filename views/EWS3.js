    var ver = 'V3.1.16';
    window.onload = start;
    var socket = io.connect();
    var ausw;

function start() {
    socket.on('chat', function (data) {
        if (document.title === "judgetool") {
            if (data.msg === 'toRoot' && parseInt(data.text) === WR_ID) {
                window.location.href = "/logout";
            }
            if (data.msg === 'WR-Info1') {
                document.getElementById(data.msg).innerHTML = data.text;
            }
            if (data.msg === 'WR-Info2') {
                document.getElementById(data.msg).innerHTML = data.text;
            }
            if (data.msg === 'aufWRwarten') {
                document.getElementById('content1').innerHTML = data.text;
            }
            if (data.msg === 'body' && parseInt(data.WR) === WR_ID) {
                document.body.innerHTML = data.HTML;
                ausw = data.ausw;
                set_events();
                return;
            }
        }
        if (document.title === "beamer") {
            if (data.seite) {
                //"beamer_seite"
            }
            if (data.kopf) {
                document.getElementById("beamer_kopf").innerHTML = data.kopf;
            }
            if (data.inhalt) {
                document.getElementById("beamer_inhalt").innerHTML = data.inhalt;
            }
            if (data.wrstatus) {
                document.getElementById("beamer_wrinfo").innerHTML = data.wrstatus;
            }
            return;
        }
        if (document.title === "moderator") {
            if (data.msg === 'mod_inhalt') {
                document.getElementById('mod_inhalt').innerHTML = data.text;
                set_events();
           }
            if (data.msg === 'mod_wrstatus') {
                document.getElementById('content1').innerHTML = data.text;
            }
        }
    });
    if (document.title === 'judgetool') {
        set_events();
        senden('get_wr_status', WR_ID);
    }
    if (document.title === 'moderator') {
        set_events();
    }
}

function set_events() {
    var t = document.getElementsByClassName("mistakes_inputs");
    for (var i = 0; i < t.length; i++) {
        t[i].setAttribute("oninput", "check_mistakes(event)");
    }
    var ev = [  "btn_wert", "wr_onclick(event)",
                "btn_leer", "wr_onclick(event)",
                "btn_NG_leer", "wr_onclick(event)",
                "btn-warning", "wr_addmistake(event)",
                "mistakes-list", "wr_delmistake(event)",
                "btn-attention", "wr_addmistake(event)",
                "obsbuttons", "obs_add(event)",
                "spalte", "wr_onclick(event)",
                "verwbutton leer", "verwarnung(event)",
                "mod_kopf", "senden_mod(event)", 
                "mod_nb", "senden_mod_zeit(event)", 
                "bs_wert", "wr_onclick(event)", 
                "bs_mist", "bs_mistake(event)",
                "weiter", "senden_sieger(event)"];

    for (var add_ev = 0; add_ev < ev.length; add_ev += 2) {
        t = document.getElementsByClassName(ev[add_ev]);
        for (i = 0; i < t.length; i++) {
            t[i].setAttribute("onclick", ev[add_ev + 1]);
        }
    }
    var s = document.getElementsByClassName("kriterium_NG");
    for (i = 0; i < s.length; i++) {
        t = document.getElementById("w" + s[i].id).value;
        if (t !== "") {
            paint_bar(s[i].children[parseFloat(t) * 2]);
        }
    }
}

function wr_delmistake(e) {
    e = e || window.event;
    var tar = e.target || e.srcElement;
    var t = tar.parentNode;
    var s = t.id.substr(13,4);
    var fl = tar.innerText.substr(1,2);
    if (tar.className === "btn-danger" || tar.className === "btn-attention") {
        var pu = document.getElementById("tfl" + s);
        pu.value = pu.value.replace(" " + tar.innerText, "");
        pu = document.getElementById("wfl" + s);
        pu.value = parseInt(pu.value) - parseInt(fl);
        tar.parentNode.removeChild(tar);
    }
}

function bs_mistake(e) {
    e = e || window.event;
    var tar = e.target || e.srcElement;
    var t = tar.parentNode;
    var i;
    var seite;
    var sum = 0;
    var te = "";
    if (tar.nodeName === "TD") {
        if (tar.className === "bs_mist") {
            tar.className = "bs_msel";
        } else {
            tar.className = "bs_mist";
        }
        for (i = 0; i <= t.childElementCount - 2; i++) {
            if (t.children[i].className === "bs_msel") {
                sum = parseFloat(sum) + parseFloat(t.children[i].getAttribute("max"));
                te = te + t.children[i].getAttribute("max") + " ";
            }
        }
        document.getElementById("w" + t.id).value = parseFloat(sum);
        document.getElementById("t" + t.id).value = te;
        seite = t.id.substring(t.id.length - 1, t.id.length); 
        add_punkte(seite);
    }
}

function add_punkte(seite) {
    var i;
    var c;
    var sum = 0;
    var va; 
    var kri = document.getElementsByClassName("bs_krit");
    for (c = 0; c < kri.length; ++c) {
        i = kri[c].id.substring(kri[c].id.length - 1, kri[c].id.length);
        va = document.getElementById("w" + kri[c].id).value;
        if (i === seite && va !== "") {
            sum = parseFloat(sum) + parseFloat(va);
        }
    }
    if (document.getElementById("wfe" + seite)) {
        sum = parseFloat(sum) - parseFloat(document.getElementById("wfe" + seite).value);
        if (sum < 0) { sum = 0; }
        document.getElementById("punktefe" + seite).innerHTML = sum;
    }
}

function obs_add(e) {
    e = e || window.event;
    var tar = e.target || e.srcElement;
    var t = tar.parentNode;
    document.getElementById("t" + t.id).value +=  " " + tar.innerText;
    document.getElementById("t" + t.id).value = document.getElementById("t" + t.id).value.trim();
    if (document.getElementById("t" + t.id).value.indexOf("-") > -1) {
        document.getElementById("t" + t.id).value = "-";
    }
    poll_mistakes();
}

function poll_mistakes() {
    var s = document.getElementsByClassName("mistakes_inputs");
    for (var i = 0; i < s.length; i++) {
        if ( check_inhalt(s[i]) ) {
            document.getElementById("absend").disabled = false;
            document.getElementById("absend").className="button_2";
        } else {
            document.getElementById("absend").disabled = true;
            document.getElementById("absend").className="button_1";
            break;
        }
    }
    function check_inhalt(mis) {
        var mistakes = ['-', 'T2' ,'T10', 'T20', 'U2', 'U10', 'U20', 'S20', 'V5','P0','A20'];
        var fld = document.getElementById('w' + mis.name.substr(1,8));
        fld.value = 0;
        mis.value = mis.value.replace(/x/gi, "-");
        var inh = mis.value.split(" ");
        for ( var f = 0; f < inh.length; f++){
            if (mistakes.indexOf(inh[f]) === -1 || inh[f] ==="" ) {
                return false;
            } else {
                fld.value = parseInt(fld.value) + to_zahl(inh[f].substr(1, 3));
            }
        }
        return true;
    }
    function to_zahl(wert) {
        if (isNaN(wert) || wert === "") {
            return 0;
        } else {
            return parseFloat(wert);
        }
    }
}

function check_mistakes(e) {
    e = e || window.event;
    var tar = e.target || e.srcElement;
    var t = tar.parentNode;
    tar.value = tar.value.trim().toUpperCase();
    poll_mistakes();
}

function senden_WRInfo() {
    var info;
    for (var s = 1; s < 3; s++) {
        if (document.getElementById('wsbs1' + s)) {
            info = '<table><tr><td width="120px">Side by Side</td><td  width="30px">' + document.getElementById('wsbs1' + s).value + '</td><td width="30px">' + document.getElementById('wsbs2' + s).value + '</td></tr>';
            info += '<tr><td>Acrobatic</td><td>' + document.getElementById('wakro' + s).value + '</td><td>&nbsp;</td></tr>';
            info += '<tr><td>Highlight</td><td>' + document.getElementById('whigh' + s).value + '</td><td>&nbsp;</td></tr></table>';
            senden('WR-Info' + s, info);
        }
    }
}

function senden_mod(e) {
    e = e || window.event;
    var tar = e.target || e.srcElement;
    senden("Moderator", tar.innerText);
}

function senden_mod_zeit(e) {
    e = e || window.event;
    var tar = e.target || e.srcElement;
    socket.emit('chat', { msg: "Moderator", text: "Paare", rnd: tar.id.substring(2, 5) });
}

function senden_sieger(e) {
    e = e || window.event;
    var tar = e.target || e.srcElement;
    var s = tar.parentNode.children
    for (var p = 0; p < s.length; p++) {
        s[p].style.backgroundColor = "#f88";
    }
    socket.emit('chat', { msg: "Moderator", text: "Sieger", rnd: tar.parentNode.id, Platz: s[0].innerText});
}

function wr_addmistake(e) {
    var pu;
    e = e || window.event;
    var tar = e.target || e.srcElement;
    var t = tar.parentNode.parentNode;
    var s = t.id.substr(8,4);
    var fl = tar.innerText.substr(1,2);

    var mist = document.getElementById("mistakes-list" + s);
    if ( mist.childElementCount < 4 ) {
        if (tar.className === "btn-attention" ) {
            if (mist.innerText.indexOf("P0") === -1 ) {
                mist.innerHTML += '<div class="btn-danger">' + tar.innerText + '</div>';
                pu = document.getElementById("tfl" + s);
                pu.value += " " + tar.innerText;
                pu = document.getElementById("wfl" + s);
                pu.value = parseInt(pu.value) + parseInt(fl);
            }
        } else {
            mist.innerHTML += '<div class="btn-danger">' + tar.innerText + '</div>';
            pu = document.getElementById("tfl" + s);
            pu.value += " " + tar.innerText;
            pu = document.getElementById("wfl" + s);
            pu.value = parseInt(pu.value) + parseInt(fl);
        }
    }
}

function wr_onclick(e) {
    e = e || window.event;
    var tar = e.target || e.srcElement;
    paint_bar(tar);
}

function paint_bar(tar) {
    var i;
    var ke;
    var t = tar.parentNode;
    var s;
    switch (ausw) {
        case "MK_":
//            if (t.type === "keyup") {
                ke = tar.value.replace(',', '.');
                if (isNaN(ke)) {
                    var we = document.getElementById(tar.id).value;
                    i = we.length - 1;
                    document.getElementById(tar.id).value = we.substr(0, i);
                    return false;
                }
                s = document.getElementsByClassName("mk_inp");
 //           }
            break;
        case "BS_":
            for (i = 0; i <= t.childElementCount - 2; i++) {
                if (i <= parseInt(tar.cellIndex)) {
                    t.children[i].className = "bs_sel";
                } else {
                    t.children[i].className = "bs_wert";
                }
            }
            document.getElementById("w" + t.id).value = parseFloat(tar.cellIndex) / 2;
            s = document.getElementsByClassName("bs_krit");
            var seite = t.id.substring(t.id.length - 1, t.id.length); 
            add_punkte(seite);  
            break;
        case "BW_":
            for (i = 0; i <= t.childElementCount - 1; i++) {
                if (i <= parseInt(tar.textContent)) {
                    t.children[i].className = "btn_sel";
                } else {
                    t.children[i].className = "btn_leer";
                }
            }
            document.getElementById("w" + t.id).value = parseInt(tar.textContent);
            s = document.getElementsByClassName("kriterium");
            break;
        case "BW_NG":
            for (i = 0; i <= t.childElementCount - 1; i++) {
                if (i <= parseInt(tar.cellIndex)) {
                    t.children[i].className = "btn_NG_sel";
                } else {
                    t.children[i].className = "btn_NG_leer";
                }
            }
            document.getElementById("w" + t.id).value = parseInt(tar.cellIndex) / 2;
            s = document.getElementsByClassName("kriterium_NG");
            break;
        case "OB_":
            i = parseInt(tar.textContent + "1");
            b = parseInt(document.getElementById("t" + t.id).innerHTML);
            i = parseFloat(document.getElementById("w" + t.id).value) + i;
            if (i < 0) { i = 0; }
            document.getElementById("w" + t.id).value = i;
            senden_WRInfo();
            s = document.getElementsByClassName("kriterium");
            return;
        case "RR_":
            for (i = 0; i <= t.childElementCount - 2; i++) {
                if (i <= parseInt(tar.cellIndex)) {
                    t.children[i].className = "btn_red";
                } else {
                    t.children[i].className = "btn_wert";
                }
            }
            var a = document.getElementById("w" + t.id).max.substr(0, 6);
            document.getElementById("w" + t.id).value = parseFloat(a) * (100 - parseInt(tar.innerText)) / 100;
            s = document.getElementsByClassName("kriterium");
            break;
        default:
            s = document.getElementsByClassName("kriterium");
    }
    for (i = 0; i < s.length; i++) {
        if (document.getElementById("w" + s[i].id).value === "NaN") {
            document.getElementById('WR-Info1').innerHTML = '<h3><b>Fehler in der Berechnung!</b></h3>';
            document.getElementById("absend").disabled = true;
            document.getElementById("absend").className = "button_1"; 
            return false;
        }
        if (document.getElementById("w"+s[i].id).value === "") {
            return false;
        }
    }
    document.getElementById('WR-Info1').innerHTML = '';
    document.getElementById("absend").disabled = false;
    document.getElementById("absend").className = "button_2"; 
}

function test_mk(te) {
    var cgivar = '';
    var ke;
    if (te.type === "keyup") {
        ke = te.target.value.replace(',', '.');
        if (isNaN(ke)) {
            var we = document.getElementById(te.target.id).value;
            var i = we.length - 1;
            document.getElementById(te.target.id).value = we.substr(0, i);
            return false;
        }
        s = document.getElementsByClassName("mk_inp");
        for (i = 0; i < s.length; i++) {
            if (document.getElementById(s[i].id).value === "NaN") {
                document.getElementById('WR-Info1').innerHTML = '<h3><b>Fehler in der Berechnung!</b></h3>';
                document.getElementById("absend").disabled = true;
                document.getElementById("absend").className = "button_1";
                return false;
            }
            if (document.getElementById(s[i].id).value === "") {
                return false;
            }
        }
        document.getElementById("absend").disabled = false;
        document.getElementById("absend").className = "button_2";
    }
}

function senden(cmd, text) {
    // Eingabefelder auslesen
    socket.emit('chat', { msg: cmd, text: text });
}

function f_send() {
//    var eingabe = window.confirm("Sicher alles gewertet?");
    var eingabe = true;
    if (eingabe === true) {
        var Jetzt = new Date();
        var Stunden = Jetzt.getHours();
        var Minuten = Jetzt.getMinutes();
        var Sekunden = Jetzt.getSeconds();
        var Vorstd = Stunden < 10 ? "0" : "";
        var Vormin = Minuten < 10 ? "_0" : "_";
        var Vorsek = Sekunden < 10 ? "_0" : "_";
        var Uhrzeit = Vorstd + Stunden + Vormin + Minuten + Vorsek + Sekunden;
        document.getElementsByName("wtim")[0].value = Uhrzeit;

        var elements = document.forms["Formular"].elements;
        var cgivar = '';
        for (var el = 0; el < elements.length; el++) {
            if (elements[el].type !== 'button') {
                cgivar += elements[el].name + '=' + elements[el].value + '&';
            }
        }
        cgivar += 'WR_ID=' + WR_ID;
        if (cgivar.indexOf("NaN") > 0) {
            document.getElementById('WR-Info1').innerHTML = '<b>Fehler in der Berechnung!</b>';
        } else {
            senden('auswerten', cgivar);
            document.getElementById("absend").disabled = true;
            document.getElementById("absend").className = "button_1";
        }
    }
    return false;
}

function chkFormular () {
    document.getElementById("absend").disabled = true;
    document.getElementById("absend").className="button_1";
    return true;
}

function p_logout() {
    window.location.href = "/logout";
}

function verwarnung(e) {
    e = e || window.event;
    var tar = e.target || e.srcElement;
    var t = tar.parentNode;
    var b = tar.className;
    switch (b) {
        case "verwbutton leer":
/*            tar.className = "verwbutton gray";
            tar.children[0].value = "0";
            break;
        case "verwbutton gray":
*/
            tar.className = "verwbutton yell";
            tar.children[0].value = "0";
            break;
        case "verwbutton yell":
            tar.className = "verwbutton red";
            tar.children[0].value = "30";
            break;
        case "verwbutton red":
            tar.className = "verwbutton black";
            tar.children[0].value = "100";
            break;
        default:
            tar.className = "verwbutton leer";
            tar.children[0].value = "";
    }
}