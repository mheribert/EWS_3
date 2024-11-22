    var ver =  'V3.2021';
    window.onload = start;
    var socket = io.connect();
    var ausw;

var drop_filled = new Object;

function start() {
    socket.on('chat', function (data) {
        if (document.title === "judgetool") {
            if (data.msg === 'judgetool' && data.text === 'toRoot') {
                window.location.href = "/logout";
            }
            if (data.msg === 'toRoot' && parseInt(data.WR) === WR_ID) {
                //                window.location.href = "/logout";
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
            if (data.msg === 'setWert' && parseInt(data.WR) === WR_ID) {
                if (data.fld === "absenden") {
                    f_send();
                } else {
                    switch (data.fld.substring(0, 3)) {
                        case "tfl":
                            if (data.val !== "") {
                                btn = document.getElementById("mistakes-list" + data.fld.substring(3));
                                btn.innerHTML = '<div class="btn-danger">' + data.val + '</div>';
                                btn = document.getElementById(data.fld);
                                btn.value = data.val;
                            }
                            break;
                        case "wfl":
                            btn = document.getElementById(data.fld);
                            btn.value = data.val / 2;
                            break;
                        default:
                            var such = ["wsbs1", "wsbs2", "wakro", "whigh"];
                            if (such.indexOf(data.fld.substring(0, data.fld.length - 1)) > -1) {
                                btn = document.getElementById(data.fld);
                                btn.value = data.val /2;
                            } else {
                                btn = document.getElementById(data.fld.substring(1));
                                btn = btn.children[data.val];
                                paint_bar(btn);
                            }
                    }
                }
            }
            if (data.msg === 'remSend' && parseInt(data.WR) === WR_ID) {
                f_send();
            }
        }
        if (document.title === 'mehrkampf') {
            var turnier = localStorage.getItem('turnier');
            if (data.turnier) {
                if (data.turnier !== turnier) {
                    localStorage.clear();
                    localStorage.setItem('turnier', data.turnier);
                    localStorage.setItem('eintraege', '');
                }
            }
            if (data.couple && parseInt(data.WR) === WR_ID) {
                var eintraege = localStorage.getItem('eintraege');
                var paar_id = data.couple.Runde + '_' + data.couple.TP_ID;
                if (eintraege.indexOf(paar_id) === -1) {
                    eintraege += paar_id + ', ';
                    var aufgabeText = { 'value': data.couple };
                    localStorage.setItem(paar_id, JSON.stringify(aufgabeText));
                    localStorage.setItem('eintraege', eintraege);
                }
            }
            if (parseInt(data.storage_load) === WR_ID) {
                fill_station();
            }
            if (parseInt(data.send_mk) === WR_ID) {
                senden_mk();
            }
            if (data.msg === 'setWert' && parseInt(data.WR) === WR_ID) {
               //   document.getElementById("mySelect").selectedIndex = "2"; 
                btn = document.getElementById(data.fld);
                btn = btn.value[data.val];
            }
            if (data.storage_clear) {
                localStorage.clear();
                localStorage.setItem('Turnier', data.turnier);
                localStorage.setItem('eintraege', '');
                fill_station();
            }
        }
        if (data.msg === 'WRlaufzeit' && parseInt(data.WR) === WR_ID) {
            socket.emit('chat', { msg: 'WR_retour', text: WR_ID });
        }
    });
    if (document.title === 'judgetool') {
        set_events();
        senden('get_wr_status', WR_ID);
    }
    if (document.title === 'moderator') {
        set_events();
    }
    if (document.title === 'mehrkampf') {
        document.getElementById("klasse").addEventListener('change', select_klasse);
        document.getElementById("station").addEventListener("change", select_station);
        s = document.getElementsByClassName("kopf_1"); 
        s[0].setAttribute('onclick', "fill_station()");
        senden('get_mk_paare', WR_ID);
        fill_station();
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
                "bs_wert", "wr_onclick(event)",
                "bs_sel", "wr_onclick(event)",
                "bs_mist", "bs_mistake(event)",
                "mk_bwert", "wr_onclick(event)",
                "mk_bsel", "wr_onclick(event)"
             ];

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
    if (document.getElementById("t" + t.id).value.split("-").length === 3) {
        document.getElementById("t" + t.id).value = "";
    }
    if (document.getElementById("t" + t.id).value.indexOf("-") > -1) {
        document.getElementById("t" + t.id).value = "-";
    }
    poll_mistakes(); "-   -"
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
        var mistakes = ['-', 'T2', 'T10', 'TF2', 'TF10','T20' ,'U2' ,'U10' ,'U20' ,'S20' ,'V5' ,'P0' ,'A20' ,'Z20'];
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

function senden_mk() {
    var eintraege = localStorage.getItem('eintraege').split(", ");
    format_zeit();
    var sende_zeit = document.getElementsByName("wtim")[0].value;
    for (var i = 0; i < eintraege.length - 1; i++) {
        var werte = localStorage.getItem('w_' + eintraege[i]);
        if (werte !== null) {
            var couple = JSON.parse(localStorage.getItem(eintraege[i]));
            var cgivar = 'MK_check=1&TP_ID1=' + couple.value.TP_ID + '&rh1=1&rt_ID=' + couple.value.RT_ID + '&' + werte + '&WR_ID=' + WR_ID + '&wtim=' + sende_zeit;

            senden('auswerten', cgivar);
        }
    }

}

function wr_addmistake(e) {
    var pu;
    e = e || window.event;
    var tar = e.target || e.srcElement;
    var t = tar.parentNode.parentNode;
    var s = t.id.substr(8, 4);
    if (tar.innerText.substr(0, 2) === "TF") {
        var fl = tar.innerText.substr(2, 2);
    } else {
        fl = tar.innerText.substr(1,2);
    }
    var mist = document.getElementById("mistakes-list" + s);
    if ( mist.childElementCount < 5 ) {
        if (tar.className === "btn-attention" ) {
            if (mist.innerText.indexOf(tar.innerText) === -1 ) {
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
//            if (tar.innerText === "A20") {
//                senden("WR-Info" + s.substring(0, 1), mist.childElementCount + " <b>A20 wurde(n) vergeben</b>")
//            }
        }
    }
}

function wr_delmistake(e) {
    e = e || window.event;
    var tar = e.target || e.srcElement;
    var t = tar.parentNode;
    var s = t.id.substr(13, 4);
    if (tar.innerText.substr(0, 2) === "TF") {
        var fl = tar.innerText.substr(2, 2);
    } else {
        fl = tar.innerText.substr(1, 2);
    }
    if (tar.className === "btn-danger" || tar.className === "btn-attention") {
        var pu = document.getElementById("tfl" + s);
        pu.value = pu.value.replace(" " + tar.innerText, "");
        pu = document.getElementById("wfl" + s);
        pu.value = parseInt(pu.value) - parseInt(fl);
        tar.parentNode.removeChild(tar);
        /*  if (tar.innerText === "A20") {
              if (t.childElementCount === 0) {
                  senden("WR-Info" + s.substring(0, 1), "alle A20 wurden gelöscht");
              } else {
                  senden("WR-Info" + s.substring(0, 1), t.childElementCount + " A20 wurde(n) vergeben");
              }
          }*/
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
    var couple;
    var erg;
    var t = tar.parentNode;
    var s;
    switch (ausw) {
        case "MK_T":
            ke = tar.value.replace(',', '.');
            if (isNaN(ke) && ke !== "-") {
                var we = ke;
                i = we.length - 1;
                tar.value = we.substr(0, i);
                return false;
            }
            tar.value=tar.value.replace('.', ',');
            couple = t.parentNode.classList[1];
            s = document.getElementsByClassName(couple);
            erg = '';
            for (i = 0; i < s.length; i++) {
                if (erg.length > 0) { erg += '&'; }
                erg += 'w' + s[i].id + '=' + s[i].childNodes[0].childNodes[0].value;
            }
            localStorage.setItem('w_' + couple, erg);
            check_mkpagefilled(); 
            return;
        case "MK_A":
            var max_pu = parseInt(t.attributes['max'].value);
            var pre = parseInt(tar.textContent);
            for (i = 0; i < 11; i++) {
                t.children[i].className = "mk_bwert";
            }
            tar.className = "mk_bsel";
            t.lastChild.firstChild.value = pre;
            couple = t.classList[1];
            s = document.getElementsByClassName(couple);
            erg = '';
            for (i = 0; i < s.length; i++) {
                if (erg.length > 0) { erg += '&'; }
                erg += 'w' + s[i].id + '=' + s[i].lastChild.firstChild.value;
            }
            check_mkpagefilled(); 
            localStorage.setItem('w_' + couple, erg);
            break;
        case "MK_B":
            var max_pu = parseInt(t.attributes['max'].value);
            var pre = parseFloat(tar.textContent);
            var g_zahl = parseInt(t.lastChild.firstChild.value) || 0;
            var k_zahl = parseFloat(t.lastChild.firstChild.value) - g_zahl || 0;
            if (tar.cellIndex < 8) {
                for (i = 0; i <= parseInt(t.attributes['max'].value); i++) {
                    t.children[i].className = "mk_bwert";
               }
                tar.className = "mk_bsel";
                g_zahl = pre;
            }
            if (tar.cellIndex > 7 || pre === max_pu) {
                if (tar.textContent === '+' || g_zahl === max_pu || tar.textContent === k_zahl.toFixed(1)) { pre = 0; }
                for (i = 1; i < 5; i++) {
                    if (i * 2 / 10 === pre) {
                        t.children[i + 8].className = "mk_bsel";
                    } else {
                        t.children[i + 8].className = "mk_bwert";
                    }
                }
                k_zahl = pre;
            }
            t.lastChild.firstChild.value = (g_zahl + k_zahl).toFixed(1).trim();
            couple = t.classList[1];
            s = document.getElementsByClassName(couple);
            erg = '';
            for (i = 0; i < s.length; i++) {
                if (erg.length > 0) {
                    erg += '&';
                }
                erg += 'w' + s[i].id + '=' + s[i].lastChild.firstChild.value
            }
            check_mkpagefilled(); 
            localStorage.setItem('w_' + couple, erg);
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
//    document.getElementById('WR-Info1').innerHTML = '';
    document.getElementById("absend").disabled = false;
    document.getElementById("absend").className = "button_2";
}

function test_mk(te) {
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
        format_zeit();

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

function format_zeit() {
    var Jetzt = new Date();
    var Stunden = Jetzt.getHours();
    var Minuten = Jetzt.getMinutes();
    var Sekunden = Jetzt.getSeconds();
    var Vorstd = Stunden < 10 ? "0" : "";
    var Vormin = Minuten < 10 ? "_0" : "_";
    var Vorsek = Sekunden < 10 ? "_0" : "_";
    var Uhrzeit = Vorstd + Stunden + Vormin + Minuten + Vorsek + Sekunden;
    document.getElementsByName("wtim")[0].value = Uhrzeit;
}

function chkFormular () {
    document.getElementById("absend").disabled = true;
    document.getElementById("absend").className="button_1";
    return true;
}

function check_mkpagefilled() {
    var s = document.forms["Formular"].elements;
    var all_filled = true;
    var isMK = true;
    format_zeit();
    for (var el = 0; el < s.length; el++) {
        if (s[el].tagName === 'INPUT' && s[el].value === '') {
            all_filled = false;
        }
        if (s[el].tagName === 'INPUT' && s[el].value === 'Absenden') {
            isMK = false;
        }
    }
    if (isMK === true) {
        var otext = s["klasse"][s["klasse"].selectedIndex];
        if (all_filled) {
            document.getElementById('couple1').style.backgroundColor = "#dfd";
            otext.textContent = '*   ' + otext.textContent.replace("*   ", "");
        } else {
            document.getElementById('couple1').style.backgroundColor = "";
            otext.textContent = otext.textContent.replace("*   ", "");
        }
        drop_filled[document.getElementById("station").value + document.getElementById("klasse").value] = all_filled;
    }
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
            tar.className = "verwbutton yell";
            tar.children[0].value = "0";
            break;
        case "bs_verw leer":
            tar.className = "bs_verw yell";
            tar.children[0].value = "3";
            break;
        case "verwbutton yell":
            tar.className = "verwbutton red";
            tar.children[0].value = "30";
            break;
        case "bs_verw yell":
            tar.className = "bs_verw red";
            tar.children[0].value = "5";
            break;
        case "verwbutton red":
            tar.className = "verwbutton black";
            tar.children[0].value = "100";
            break;
        case "bs_verw red":
            tar.className = "bs_verw leer";
            tar.children[0].value = "0";
            break;
        case "verwbutton black":
            tar.className = "verwbutton leer";
            tar.children[0].value = "0";
            break;
        default:
            tar.classList[1] = "leer";
            tar.children[0].value = "";
    }
}

function fill_station() {
    var station = new Object;
    station[0] = '---';
    var eintraege = localStorage.getItem('eintraege').split(", ");
    for (var i = 0; i < eintraege.length - 1; i++) {
        var couple = JSON.parse(localStorage.getItem(eintraege[i]));
        station[couple.value.Runde] = couple.value.T_Text;
        drop_filled[couple.value.Runde + couple.value.st_kl] = false;
    }
    fill_select("station", station);
    document.getElementById("klasse").innerHTML = '<option value="0">---</option>';
    document.getElementById("wertungen").innerHTML = '<td class="main" height="300px"></td>';
    document.getElementById("station").focus();
}

function select_station() {
    var klasse = new Object;
    klasse[0] = '---';
    var eintraege = localStorage.getItem('eintraege').split(", ");
    for (var i = 0; i < eintraege.length - 1; i++) {
        var couple = JSON.parse(localStorage.getItem(eintraege[i]));
        klasse[couple.value.st_kl] = couple.value.startkl;
    }
    fill_select("klasse", klasse);
    document.getElementById("wertungen").innerHTML = '<td class="main" height="300px"></td>';
    document.getElementById("klasse").focus();
}

function fill_select(dropdown, werte) {
    var dr = document.getElementById(dropdown);
    dr.innerHTML = '';
    for (var i in werte) {
        var option = document.createElement("option");
        option.text = werte[i];
        option.value = i;
        if (dropdown !== "station") {
            if (drop_filled[document.getElementById("station").value + option.value] === true) {
                option.text = '*   ' + option.text;
            }
        }
        dr.add(option);
    }
}

function select_klasse() {
    var paare = new Object;
    var HTML_Seite;
    var menu = document.getElementById("klasse");
    var kl = menu.options[menu.selectedIndex].value;
    menu = document.getElementById("station");
    var rde = menu.options[menu.selectedIndex].value;
    paare[0] = '---';
    var eintraege = localStorage.getItem('eintraege').split(", ");
    HTML_Seite = '<td align="center" id="couple1" height="300px"><table align="center" border="0" cellpadding="0" cellspacing="0">' + '\r\n';
    for (var i = 0; i < eintraege.length - 1; i++) {
        var couple = JSON.parse(localStorage.getItem(eintraege[i]));
        if (couple.value.st_kl === kl && couple.value.Runde === rde) {
            HTML_Seite += '<tr><td class="sel_paare" colspan ="21">' + couple.value.Startnr + '   ' + couple.value.Dame + ' - ' + couple.value.Herr + '</td></tr>';
            HTML_Seite += select_paare(couple.value);
        }
    }
    HTML_Seite += '</table></td>';
    document.getElementById("wertungen").innerHTML = HTML_Seite;
    set_events();
    check_mkpagefilled();
}

function select_paare(val) {
    var sei = 1;
    var focus_to;
    var HTML_Seite;
    var trunde = val.Runde.substr(0, 4);
    var paar = val.Runde + '_' + val.TP_ID;
    var werte = hole_eintrag('w_' + paar);
    if (wr_func === "MA") {
        if (trunde === "MK_3" || trunde === "MK_4") {
            HTML_Seite = make_inpMK('mk_td' + sei, 10, 'Dame', true, werte['wmk_td' + sei], paar) + '\r\n';
            HTML_Seite += '<tr><td height="20"></td></tr>' + '\r\n';
            HTML_Seite += make_inpMK('mk_th' + sei, 10, 'Herr', true, werte['wmk_th' + sei], paar) + '\r\n';
            HTML_Seite += '<tr><td height="200"></td></tr>' + '\r\n';
            ausw = "MK_A";
        } else {
            HTML_Seite = '<tr><td height="10"></td></tr>' + '\r\n';
            HTML_Seite += make_inpMKText('mk_td' + sei, 0, "Dame", werte['wmk_td' + sei], paar) + '\r\n';
            HTML_Seite += '<tr><td height="20"></td></tr>' + '\r\n';
            HTML_Seite += make_inpMKText('mk_th' + sei, 0, "Herr", werte['wmk_th' + sei], paar) + '\r\n';
            HTML_Seite += '<tr><td height="150"></td></tr>' + '\r\n';
            ausw = "MK_T";
        }
        focus_to = 'wmk_td' + sei;
    } else {   //  MB
        if (trunde === "MK_3" || trunde === "MK_4") {
            HTML_Seite = make_inpMK('mk_td' + sei, 7, 'Dame Technik & Haltung', false, werte['wmk_td' + sei], paar) + '\r\n';
            HTML_Seite += make_inpMK('mk_dd' + sei, 3, 'Dame Dynamik', false, werte['wmk_dd' + sei], paar) + '\r\n';
            HTML_Seite += '<tr><td height="30"></td></tr>' + '\r\n';
            HTML_Seite += make_inpMK('mk_th' + sei, 7, 'Herr Technik & Haltung', false, werte['wmk_th' + sei], paar) + '\r\n';
            HTML_Seite += make_inpMK('mk_dh' + sei, 3, 'Herr Dynamik', false, werte['wmk_dh' + sei], paar) + '\r\n';
            HTML_Seite += '<tr><td height="100"></td></tr>' + '\r\n';
            ausw = "MK_B";
        } else {
            HTML_Seite = '<tr><td height="270">Kein Einsatz</td></tr>' + '\r\n';
        }
        focus_to = 'station';
    }
    return HTML_Seite;
}

function hole_eintrag(couple) {
    var cat;
    var wert = new Object();
    var i;
    var vorh = localStorage.getItem(couple);
    if (vorh !== null) {
        vorh = vorh.split("&");
        for (i = 0; i < vorh.length; i++) {
            cat = vorh[i].split("=");
            wert[cat[0]] = cat[1];
        }
    }
    return wert;
}

function make_bs_inp(fName, max, aName, ganz, pre, paar) {
    var inp;
    var b_class;
    if (pre === undefined) {
        pre = -1;
    }
    inp = '<tr class="bs_head"><td colspan="20">' + aName + '</td></tr>';
    inp += '<tr class="bs_krit ' + paar + '" id="' + fName + '" max="' + max + '">';
    for (var t = 0; t < max * 2 + 1; t++) {
        if (pre * 2 < t) {
            b_class = "bs_wert";
        } else {
            b_class = "bs_sel";
        }
        if (t % 2) {
            if (ganz) {
                inp += '<td style="width:40px; height:40px; visibility: hidden;" class="' + b_class + '">' + '-' + '</td>';
            } else {
                inp += '<td style="width:40px; height:40px;" class="' + b_class + '">' + '-' + '</td>';
            }
        } else {
            inp += '<td style="width:40px; height:40px;" class="' + b_class + '">' + t / 2 + '</td>';
        }
    }
    inp += '<input name="w' + fName + '" id="w' + fName + '" value="' + pre + ' " type="hidden"></tr>';

    return inp;
}

function make_inpMK(fName, max, aName, ganz, pre, paar) {
    var inp;
    var b_class;
    var g_zahl = parseInt(pre);
    var k_zahl = (parseFloat(pre) - g_zahl).toFixed(1);
    inp = '<tr class="bs_head"><td><table align="center"><tr><td colspan="20">' + aName + '</td></tr>';
    inp += '<tr class="bs_krit ' + paar + '" id="' + fName + '" max="' + max + '">';
    if (max === 10) {
        for (var t = 0; t < max +1 ; t++) {
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

function make_inpMKText(fName, max, aName, pre, paar) {
    var inp;
    if (pre === undefined) {
        pre = "";
    }
    inp = '<tr><td><table align="center"><tr><td colspan="20">' + aName + '</td></tr>';
    inp += '<tr class="mk_inp ' + paar + '" id="' + fName + '"><td><input class="mk_fld" id="w' + fName + '" name="w' + fName + '" value="' + pre + '" autocomplete="off" onkeyup="wr_onclick(event)"></td></tr>';
    inp += '</table></td></tr>';

    return inp;
}

function reise_summe() {
    var km = document.getElementById("r_km").value;
    var summe;
    if (km > 300) {
        km -= 300;
        summe = (300 * 0.3 + km * 0.15) * 2;
    } else {
        summe = km * 0.15 * 2;
    }
    summe += (parseInt(document.getElementById("r_kbahn").value) || 0);

    summe += (parseInt(document.getElementById("r_pausch14").value) || 0) * 14;
    summe += (parseInt(document.getElementById("r_pausch28").value) || 0) * 28;

    summe = parseFloat(summe) - (parseInt(document.getElementById("r_frueh").value) || 0) * 5.60;
    summe = parseFloat(summe) - (parseInt(document.getElementById("r_essen").value) || 0) * 11.20;
    summe = parseFloat(summe) - (parseInt(document.getElementById("r_abend").value) || 0) * 11.20;

    summe += (parseInt(document.getElementById("r_uekosten").value) || 0);
    summe += (parseInt(document.getElementById("r_khono").value) || 0);

    document.getElementById("r_summe").value = summe.toFixed(2);
}

function reise_send() {
    format_zeit();

    var elements = document.forms["Formular"].elements;
    var cgivars = cgivars = 'WR_ID=' + WR_ID;
    for (var el = 0; el < elements.length; el++) {
        if (elements[el].type !== 'button') {
            cgivars +=  '&' + elements[el].id + '=' + elements[el].value;
        }
    }
    reise_summe();
    socket.emit('chat', { msg: 'reise_schreib', text: cgivars });
    location.reload();
}

function getPaging(e) {
    e = e || window.event;
    var tar = e.target || e.srcElement;

    var sheet = document.styleSheets[0];
    if (sheet.rules[47].style.display === "block") {
        sheet.rules[47].style.display = "none";
    } else {
        sheet.rules[47].style.display = "block";
    }
    switch (tar.id) {
        case "li_zeit":
            window.open("zeitplan.html");
            break;
        case "li_reise":
            senden("reise_fill", WR_ID);
            break;
        case "li_back":
            location.reload();
            break;
    }
}
