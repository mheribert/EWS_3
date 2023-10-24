var ver = 'V3.2018';
window.onload = start;
var socket = io.connect();

function start() {
    socket.on('chat', function (data) {
        if (data.msg === 'beamer' && document.title === 'beamer') {
            // beamer_bild, beamer_kopf, beamer_inhalt, beamer_seite, beamer_minute
            if (data.bereich !== undefined) {
                document.getElementById(data.bereich).innerHTML = data.cont;
            }
            if (data.cont === 'reload') {
                location.reload(true);
            }
            return;
        }
        if (data.msg === 'beamer2' && document.title === 'beamer2') {
            if (data.bereich !== undefined) {
                document.getElementById(data.bereich).innerHTML = data.cont;
            }
            return;
        }
        if (data.msg === 'beamer3' && document.title === 'beamer3') {
            if (data.bereich !== undefined) {
                document.getElementById(data.bereich).innerHTML = data.cont;
            }
            return;
        }
        if (document.title === "moderator") {
            if (data.msg === 'mod_inhalt') {
                document.getElementById('mod_inhalt').innerHTML = data.text;
                set_events();
                window.scrollTo(0, 0);
            }
            if (data.msg === 'mod_wrstatus') {
                document.getElementById('content1').innerHTML = data.text;
            }
        }
 
    });
    if (document.title === 'moderator') {
        set_events();
    }
}

function set_events() {
    var ev = [
        "mod_kopf", "senden_mod(event)",
        "mod_nb", "senden_mod_zeit(event)",
        "mod_ns", "senden_mod_sieger(event)",
        "weiter", "senden_sieger(event)"
    ];

    for (var add_ev = 0; add_ev < ev.length; add_ev += 2) {
        t = document.getElementsByClassName(ev[add_ev]);
        for (i = 0; i < t.length; i++) {
            t[i].setAttribute("onclick", ev[add_ev + 1]);
        }
    }
}

function senden_mod(e) {
    e = e || window.event;
    var tar = e.target || e.srcElement;
    socket.emit('chat', { msg: "Moderator", text: tar.innerText });
}

function senden_mod_zeit(e) {
    e = e || window.event;
    var tar = e.target || e.srcElement;
    socket.emit('chat', { msg: "Moderator", text: "Paare", rnd: tar.id.substring(2, 5) });
}

function senden_mod_sieger(e) {
    e = e || window.event;
    var tar = e.target || e.srcElement;
    socket.emit('chat', { msg: "Moderator", text: "start_Sieger", rnd: tar.id.substring(2, 5), rtid: tar.getAttribute("rtid") });
}

function senden_sieger(e) {
    e = e || window.event;
    var tar = e.target || e.srcElement;
    if (tar.cellIndex === 0) { 
        var s = tar.parentNode.children;
        for (var p = 0; p < s.length; p++) {
            s[p].style.backgroundColor = "#f88";
        }
        socket.emit('chat', { msg: "Moderator", text: "Sieger", rnd: tar.parentNode.id, Platz: s[0].innerText });
    }
}
