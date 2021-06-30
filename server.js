var ver            = 'V3.2.00';
var express        = require('express');
var app		       = express();
var server         = require('http').createServer(app);
var io             = require('socket.io').listen(server);
var bodyParser     = require('body-parser');
var session 	   = require('express-session');
var conf           = require('./config.json');
var ADODB          = require('node-adodb'), colors = require('colors/safe');
var HTML_auswerten = require('./HTML_auswerten');
var HTML_erstellen = require('./HTML_erstellen');
var HTML_moderator = require('./HTML_moderator');
var HTML_beamer    = require('./HTML_beamer');
var fs             = require('fs');
//app.set('views', __dirname + '/views');
app.engine('html', require('ejs').renderFile);

// app.use(session({ secret: 'ssshhhhh', saveUninitialized: true, resave: true, cookie: { maxAge: 6000 } }));
var sec = Math.random().toString().substring(3);
app.use(session({ secret: sec, saveUninitialized: true, resave: true}));
sec = '';
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({extended: true}));
app.use(express.static(conf.pfad + '\\webserver\\views'));

var connection = ADODB.open('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + conf.pfad + conf.db + ';');

var sess;
var wertungsrichter = new Object();
var anz_wr = 0;
var observer = new Object();
var anz_obs;
var wr_status;
var Index_Seite = '<!DOCTYPE html><head><title>Judgetool</title></head><body><p style="font-size:50pt;"  align="center"><br><br>Turnier startet bald.</p></body></html>';
var HTML_Seite;
var runden_info = new Object();
var wertungen = new Object();
var akrobatiken = new Object();
var WR_Info1 = new Object();
var WR_Info2 = new Object();
var moderator_inhalt;
var title;
var turnier_nr;
var runde = 1;          // welche Tanzrunde gerade läuft
var rundenende = false; // test on alle wertungen da sind
var last_rt_id;
var new_guidelines = false;

server.listen(conf.port);

app.post('/login',function(req,res){
    connection
//        .query('SELECT * FROM (SELECT WR_kenn, WR_ID FROM wert_richter UNION SELECT PROP_VALUE, PROP_KEY FROM Properties WHERE val(PROP_KEY) > 9999) WHERE WR_ID="' + req.body.wr_id + '";')
        .query('SELECT * FROM wert_richter WHERE WR_ID=' + req.body.wr_id)
        .on('done', function (data) {
            if (req.body.passwort === data[0].WR_kenn) {
                sess = req.session;
                sess.user_id = req.body.wr_id;
                res.redirect('/judge');
            } else {
                res.redirect('/');
            }
        });
});

app.get('/judge', function (req, res) {
    sess = req.session;
    if (sess.user_id) {
        var wr_name = wertungsrichter[sess.user_id].WR_Nachname.substr(0, 1) + wertungsrichter[sess.user_id].WR_Vorname || "";
        HTML_erstellen.blankPage(0, wr_name, sess.user_id, runden_info, res);
    } else { // kein angemeldeter User
        res.redirect('/');
    }
});

app.get('/', function (req, res) {
	sess=req.session;
	if(sess.user_id) {
		res.redirect('/judge');
	} else {
		res.send(Index_Seite);
	}
});

app.get("/cgi-bin", function (req, res) {
    res.redirect('/');
});

app.get("/beamer", function (req, res) {
    res.send(HTML_beamer.beamer_seite());
    setTimeout(function () {
        io.emit('chat', HTML_beamer.inhalt());
    }, 500);
});

app.get("/moderator", function (req, res) {
    res.send(HTML_moderator.mod_seite());
    setTimeout(function () {
        io.emit('chat', { msg: 'mod_inhalt', text: HTML_moderator.inhalt() });
    }, 500);
});

app.get("/login", function (req, res) {
    res.redirect('/');
});

app.get('/logout', function(req, res) {
	req.session.destroy(function(err) {
		if(err){
			console.log(err);
		} else {
			res.redirect('/');
		}
	});
});
 
app.get('/hand', function (req, res) {
    var n;
    var i;
    switch (req.query.msg) {
        case "observer_starten":
            wertungen = new Object;
            connection = ADODB.open('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + conf.pfad + req.query.mdb + '_TDaten.mdb;');
            connection
                .query('SELECT * FROM turnier;')
                .on('done', function (data) {
                    title = data[0].Turnier_Name;
                    turnier_nr = "T" + data[0].Turnier_Nummer + "_RT";
                    connection										// wertungsrichter auffrischen
                        .query('SELECT * FROM View_Rundenablauf WHERE RT_ID =' + req.query.text + ' ORDER BY Rundennummer, Startnr;')
                        .on('done', function (data) {
                            runden_info = data;		// Rundeninfo laden
                            last_rt_id = runden_info[0].RT_ID;
                            runde = 1;
                            var HTML_Kopf = runden_info[0].Turnier_Name + '<br>' + runden_info[0].Tanzrunde_Text;
                            connection
                                .query('SELECT * FROM wert_richter ORDER BY WR_Kuerzel;')
                                .on('done', function (data) {
                                    var wr_func = "";
                                    var anz = 0;
                                    anz_wr = 0;
                                    anz_obs = 0;
                                    var rd;
                                    var trunde;
                                    var st_kl;
                                    var v_mk;
                                    if (typeof runden_info[0] === "undefined") {
                                        st_kl = "RR_A";
                                        trunde = "VR";
                                        rd = "Vor_r";
                                    } else {
                                        st_kl = runden_info[0].Startklasse;
                                        trunde = runden_info[0].RundeArt;
                                        rd = runden_info[0].Runde;
                                        v_mk = rd.substr(0, 4);
                                    }
                                    var erster = false;
                                    observer = new Object;
                                    wertungsrichter = new Object();
                                    for (i in data) {
                                        wertungsrichter[data[i].WR_ID] = data[i];
                                        wr_func = wertungsrichter[data[i].WR_ID].WR_func;
                                        if (wr_func === "Ob") {
                                            if (trunde === 'ER' && erster === true && runden_info[0].Runde !== 'Semi') {
                                                wertungsrichter[data[i].WR_ID].WR_func = "";
                                                wr_func = "";
                                                wertungsrichter[data[i].WR_ID].WR_status = "";
                                            }
                                            erster = true;
                                        }   // bei A/B Ft + S alle Ak-WR raus 
                                        if (wr_func === "Ak" && (st_kl === "RR_S" || rd.indexOf("_Fu") > 0)) {
                                            wertungsrichter[data[i].WR_ID].WR_func = "";
                                            wr_func = "";
                                            wertungsrichter[data[i].WR_ID].WR_status = "";
                                        }
                                        // Mehrkampf Stationen
                                        if (v_mk === "MK_1" || v_mk === "MK_2" || v_mk === "MK_3" || v_mk === "MK_4") {
                                            // alle normalen WR raus
                                            if (rd.substr(0, 3) === "MK_" && wr_func !== "MA" && wr_func !== "MB" && wr_func !== "Ob") {
                                                wertungsrichter[data[i].WR_ID].WR_func = "";
                                                wr_func = "";
                                                wertungsrichter[data[i].WR_ID].WR_status = "";
                                            }
                                            // bei MK 1 und 2 alle B-WR raus
                                            if ((v_mk === "MK_1" || v_mk === "MK_2") && wr_func === "MB") {
                                                wertungsrichter[data[i].WR_ID].WR_func = "";
                                                wr_func = "";
                                                wertungsrichter[data[i].WR_ID].WR_status = "";
                                            }
                                        } else {
                                            // Mehrkampf Tanz und allgemein Tanz alle mk_WR raus
                                            if (wr_func === "MA" || wr_func === "MB") {
                                                wertungsrichter[data[i].WR_ID].WR_func = "";
                                                wr_func = "";
                                                wertungsrichter[data[i].WR_ID].WR_status = "";
                                            }   
                                        }
                                        // zähle WR
                                        if (wr_func !== "") {
                                            anz_wr++;
                                        }
                                        if (wr_func === "Ob") {
                                            observer[data[i].WR_ID] = anz_obs;
                                            anz_obs++;
                                        }
                                    }
                                    Index_Seite = HTML_erstellen.wr_login(wertungsrichter, title);

                                    for (var w in wertungsrichter) {
                                        if (wertungsrichter[w].WR_status !== "") {
                                            if (wertungsrichter[w].WR_func === "Ob") {
                                                wertungsrichter[w].WR_status = 'runde';
                                            } else {
                                                wertungsrichter[w].WR_status = 'start';
                                            }
                                        }
                                        verteilen(wertungsrichter[w].WR_ID);
                                    }
                                    io.emit('chat', { msg: 'beamer', kopf: HTML_Kopf, inhalt: '<tr><td>&nbsp;</td></tr>' });
                                    HTML_moderator.runde(io, runden_info, runde);
                                });

                            console.log(runden_info[0].Tanzrunde_Text + " gestartet");
                            connection
                                .query('SELECT [NR#], Akrobatik, Langtext, ' + runden_info[0].Startklasse + ' FROM Akrobatiken WHERE ' + runden_info[0].Startklasse + ' <> "";')
                                .on('done', function (data) {
                                    for (i in data) {
                                        akrobatiken[data[i].Akrobatik] = data[i];
                                    }
                                });
                        res.send('gestartet ' + req.query.text);
                        });
                });
            break;
        case "storage_clear":
            io.sockets.emit('chat', { msg: 'mehrkampf', storage_clear : ' ' });
            res.send(req.query.msg + req.query.text);
            break;
        case "storage_send":
            connection
                .query('SELECT * FROM View_Rundenablauf WHERE Runde LIKE "MK_%" ORDER BY Startklasse, Startnr;')
                //        .query('SELECT * FROM View_Rundenablauf WHERE Runde = "MK_3_BOT" ORDER BY Startklasse, Rundennummer, Startnr;')
                .on('done', function (data) {
                    var s;
                    io.sockets.emit('chat', { msg: 'mehrkampf', turnier: data[0].RT_file.substr(0, 8) });
                    for (var i in data) {
                        var mehrkampf = new Object;
                        if (data[i].Runde.substring(0, 4) !== "MK_5") {
                            mehrkampf.RT_ID = data[i].RT_ID;
                            mehrkampf.TP_ID = data[i].TP_ID;
                            mehrkampf.st_kl = data[i].Startklasse;
                            s = data[i].Tanzrunde_Text.indexOf(" - ");
                            mehrkampf.T_Text = data[i].Tanzrunde_Text.substring(0, s);
                            mehrkampf.startkl = data[i].Tanzrunde_Text.substring(s + 3);
                            mehrkampf.Startnr = data[i].Startnr;
                            mehrkampf.Runde = data[i].Runde;
                            mehrkampf.Herr = data[i].Herr;
                            mehrkampf.Dame = data[i].Dame;
                            io.sockets.emit('chat', { msg: 'mehrkampf', WR: '4', couple: mehrkampf });
                        }
                    }
                    console.log("daten geschrieben");
                    runden_info = data;		// Rundeninfo laden
                });
            res.send(req.query.msg + req.query.text);
            break;
        case "storage_load":
            io.sockets.emit('chat', { msg: 'mehrkampf', storage_load: ' ' });
            res.send(req.query.msg + req.query.text);
            break;
        case "eingriff":
            if (req.query.text === 'runde_mi') {
                runde--;
                io.emit('chat', { msg: 'judgetool', text: 'aufWRwartenweiter' });
            }
            if (req.query.text === 'runde_pl') {
                runde++;
                io.emit('chat', { msg: 'judgetool', text: 'aufWRwartenweiter' });
            }
            res.send(req.query.msg + req.query.text);
            break;
        case "Runde_starten":
            if (runden_info[0].Tanzrunde_MAX >= runde) {
                for (i in wertungsrichter) {
                    if (wertungsrichter[i].WR_func !== "") {
                        wertungsrichter[i].WR_status = 'werten';
                        verteilen(i);
                    }
                }
            } else {
                for (i in wertungsrichter) {
                    if (wertungsrichter[i].WR_func !== "") {
                        wertungsrichter[i].WR_status = "ende";
                        verteilen(i);
                    }
                }
            }
            io.sockets.emit('chat', { zeit: new Date(), msg: 'judgetool', text: 'toRoot' });
            HTML_beamer.beamer_runde(io, runden_info, runde);
            HTML_moderator.runde(io, runden_info, runde);
            res.send("Runde auswerten");
            break;
        case "Runde_auswerten":
            if (rundenende === true) {
                HTML_auswerten.berechne_punkte(wertungen, runden_info, runde, wertungsrichter, conf.pfad + turnier_nr + runden_info[0].RT_ID + '.txt');
                HTML_beamer.beamer_ranking(io, runden_info, runde);
                HTML_moderator.runde(io, runden_info, runde);
                runde++;
                if (runde > runden_info[0].Tanzrunde_MAX) {
                    for (i in wertungsrichter) {
                        if (wertungsrichter[i].WR_func !== "") {
                            wertungsrichter[i].WR_status = 'ende';
                        }
                    }
                    io.sockets.emit('chat', { zeit: new Date(), msg: 'judgetool', text: 'aufWRwartenweiter' });
                }
            }
            res.send("Runde starten");
            break;
        case "Runde_beenden":
            io.emit('chat', { msg: 'beamer', kopf: ' ', inhalt: ' ' });
            HTML_moderator.zeitplan(io, connection);
            to_Root();
            res.send("runde beendet");
            break;

        case "WR-Info1":
            io.emit('chat', { msg: req.query.msg, text: req.query.text });
            res.send("WR gesendet");
            break;

        case "WR-Info2":
            io.emit('chat', { msg: req.query.msg, text: req.query.text });
            res.send("WR gesendet");
            break;

        case "nochmal_starten":
            for (n in runden_info) {
                if (runden_info[n].TP_ID.toString() === req.query.text) {
                    runden_info[runden_info.length] = uebertrage(runden_info[n]);
                    runden_info[n].nochmal = true;
                    runden_info[runden_info.length - 1].PpR = 1;
                    runden_info[runden_info.length - 1].Rundennummer = runden_info[n].Tanzrunde_MAX + 1;
                }
            }
            for (i in runden_info) {
                runden_info[i].Tanzrunde_MAX++;
            }
            for (i in wertungen) {
                for (n in wertungen[i]) {
                    if (n === req.query.text) {
                        delete wertungen[i][n];
                    }
                }
            }
            res.send("eingetragen");
            break;

        case "wr_lesen":
            wr_lesen();
            io.emit('chat', { msg: 'judgetool', text: 'aufWRwartenweiter' });
            res.send("wr gelesen");
            break;

        case "nochmal werten":
            for (n in wertungsrichter) {
                if ((n === req.query.text || req.query.text === "Alle") && wertungsrichter[n].WR_func !== "") {
                    wertungsrichter[n].WR_status = "werten";
                    verteilen(n);
                    runde_zurucksetzen(runde, n);
                }
            }
            if (req.query.text === "Alle") {
                runde_zurucksetzen(runde, 0);
            }
            refresh_wait();
            res.send("alle werten");
            break;

        case "status_wr":
            for (n in wertungsrichter) {
                console.log(wertungsrichter[n].WR_Nachname + '    ' + wertungsrichter[n].WR_status);
            }
            console.log('__________________');
            res.send("log geschrieben");
            break;
        case "wiederherstellen":
            res.send(req.query.msg + req.query.text);
            runde_wiederherstellen(runden_info);
            break;
        case "beamer_zeitplan":
            res.send(req.query.msg + req.query.text);
            HTML_beamer.beamer_zeitplan(io, connection, req.query.text);
            break;
        case "beamer_ranking":
            res.send(req.query.msg + req.query.text);
            HTML_beamer.beamer_ranking(io, runden_info, runde);
            break;
        case "beamer_runde":
            res.send(req.query.msg + req.query.text);
            HTML_beamer.beamer_runde(io, runden_info, runde);
            break;
        case "beamer_siegerehrung":
            res.send(req.query.msg + req.query.text);
            connection = ADODB.open('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + conf.pfad + req.query.mdb + '_TDaten.mdb;');
            HTML_beamer.beamer_siegerehrung(io, connection, req.query.text, req.query.Platz);
            HTML_moderator.siegerehrung(io, connection, req.query.text);
            break;
        case "beamer":
            res.send(req.query.msg);
            io.emit('chat', { msg: req.query.msg, seite: req.query.seite, inhalt: req.query.inhalt, kopf: req.query.kopf, wr_info: req.query.wr_info });
            break;
        case "moderator_vorstellung":
            connection = ADODB.open('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + conf.pfad + req.query.mdb + '_TDaten.mdb;');
            HTML_moderator.vorstellung(io, connection, req.query.text);
             res.send(req.query.msg + req.query.text);
            break;
        case "WR-Simulate":
            var cgivar = "";
            for (var el in req.query) {
                if (el !== 'msg') {
                    cgivar += '&' + el + '=' + req.query[el];
                }
            }
            res.send(req.query.msg + cgivar);
            io.emit('chat', { msg: req.query.msg, text: cgivar.substr(1) });
            break;
        default:
            res.send(req.query.msg + req.query.text);
            io.emit('chat', { msg: req.query.msg, text: req.query.text });
            io.emit('chat', { msg: 'beamer', kopf: '', inhalt: '' });
            HTML_moderator.runde(io, runden_info, runde);
    }
    function uebertrage(quelle) {
        var temp = new Object();
        for (var x in quelle) {
            temp[x] = quelle[x];
        }
        return temp;
    }

});

function runde_zurucksetzen(rd, WR_ID) {
    for (var i in wertungen) {
        for (var n in wertungen[i]) {
            if (wertungen[i][n].Runde === rd && WR_ID === i) {
                delete wertungen[i][n];
            }
        }
    }
}

function runde_wiederherstellen(runden_info) {
    var restore;
    if (typeof runden_info[0] === "undefined") {
        console.log('Runde noch nicht gestartet');
        return;
    }
    var contents = fs.readFileSync(conf.pfad + turnier_nr + last_rt_id + '_raw.txt', 'utf8');
    wertungen = new Object();
    var einzel = contents.split('\r\n');
    for (restore in einzel) {
        if (restore < 9) {                              // einzel[i] !== "") {
            var wr = einzel[restore].split(';');
            auswerten(wr[2].replace(/PR_ID/g, "TP_ID"));
            if (einzel[restore] === einzel[restore + 1]) {
                restore++;
            }
        }
    }
    console.log('Runde wiederhergestellt');
}

function insert_differences(body, wertungen) {
    var s;
    for (var i in body) {
        if (i.substr(0, 3).indexOf("tfl") !== -1) {
            s = i.replace("ak", "");
            s = s.substr(3, 1);
            wertungen[body.WR_ID][body["TP_ID" + s]][i] = body[i];
            wertungen[body.WR_ID][body["TP_ID" + s]]["korr"] = "Ok";
            wertungen[body.WR_ID][body["TP_ID" + s]]["w" + i.substr(1, 8)] = body["w" + i.substr(1, 8)];
        }
    }
}

// Websocket
io.sockets.on('connection', function (socket) {
    socket.on('chat', function (data) {
//        console.log('msg: ' + data.msg + '   text:' + data.text);
        switch (data.msg) {
            case "nothing":
                break;
            case "Runde_starten":
                for (var i in wertungsrichter) {
                    if (wertungsrichter[i].WR_func !== "") {
                        wertungsrichter[i].WR_status = 'werten';
                        verteilen(i);
                    }
                }
                io.sockets.emit('chat', { zeit: new Date(), msg: 'aufWRwartenweiter', text: '' });
                break;

            case "WR-Info1":
                WR_Info1 = { zeit: new Date(), msg: data.msg, text: data.text };
                io.sockets.emit('chat', WR_Info1);		// dieser Text wird an alle anderen WR gesendet
                break;

            case "WR-Info2":
                WR_Info2 = { zeit: new Date(), msg: data.msg, text: data.text };   	
                io.sockets.emit('chat', WR_Info2);		// dieser Text wird an alle anderen WR gesendet
                break;

            case "Moderator":
                switch (data.text) {
                    case "Runde":
                        HTML_moderator.runde(io, runden_info, runde);
                        break;

                    case "WR":
                        HTML_moderator.wr(io, wertungsrichter);
                        break;

                    case "Paare":
                        HTML_moderator.vorstellung(io, connection, data.rnd);
                        break;

                    case "Zeitplan":
                        if (typeof last_rt_id === "undefined") {
                            HTML_moderator.zeitplan(io, connection);
                        } else {
                            HTML_moderator.zeitplan(io, connection, last_rt_id);
                        }
                        break;
                    case "Vorst_1":
                    case "Vorst_2":
                        HTML_moderator.vorstellung(io, connection);
                        break;
                    case "Sieger":
                        HTML_beamer.beamer_siegerehrung(io, connection, data.rnd, data.Platz);
                        break;
                }
                break;
            case "nochmal werten":
                for (var n in wertungsrichter) {
                    if ((n === data.text.toString() || data.text.toString() === "Alle") && wertungsrichter[n].WR_func !== "") {
                        wertungsrichter[n].WR_status = "werten";
                        verteilen(data.text);
                    }
                }
                refresh_wait();
                break;
            case "auswerten":
                auswerten(data.text);
                break;
            case "get_wr_status":
                verteilen(data.text);
                break;
          default:
                io.sockets.emit('chat', { zeit: new Date(), msg: data.msg, text: data.text });		// dieser Text wird an alle anderen WR gesendet
                break;
        }
    });
    socket.on('disconnect', function () {
//       console.log('disconnect : ' + socket.id);
    });
    socket.on('connect', function (client) {
        console.log(client);
    });
});

function verteilen(WR_ID) {
    var rd_ind = 0;
    for (i = 0; i < runden_info.length; i++) {
        if (runden_info[i].Rundennummer < runde) {
            rd_ind++;
        }
    }
    var i;
    var wr_name = wertungsrichter[WR_ID].WR_Nachname.substr(0, 1) + wertungsrichter[WR_ID].WR_Vorname || "";
    switch (wertungsrichter[WR_ID].WR_status) {
        case "start":		// allgemeines warten vor Rundenstart
            HTML_erstellen.wait(rd_ind, runden_info, '<div class="wertung_offen" align="center">Runde beginnt bald</div>', wr_name, WR_ID, io);
            break;

        case "ende":		// allgemeines warten bei Rundenende
            HTML_erstellen.wait(rd_ind, runden_info, '<div class="wertung_offen" align="center">Rundenende<br>die neue Runde startet bald.</div>', wr_name, WR_ID, io);
            break;

        case "runde":		// Observer warten für Rundenbeginn
            if (typeof runden_info[0] === "undefined") {
                HTML_erstellen.wait(rd_ind, runden_info, '<div class="wertung_offen" align="center">Runde beginnt bald</div>', wr_name, WR_ID, io);
            } else {
                if (wertungsrichter[WR_ID].WR_func === "Ob") {
                    if (observer[WR_ID] === 1 && runden_info[rd_ind].PpR === 1) {
                        //bei zwei Ob und einem Paar bei 2 PpR
                        HTML_erstellen.wait(rd_ind, runden_info, '<div class="wertung_offen" onclick="return p_logout()">kein Einsatz in dieser Runde</div>', wr_name, WR_ID, io);
                        setTimeout(abmelden, 3000, WR_ID);
                    } else {
                        var obs_text = '<div>n&auml;chste Runde: ' + runden_info[rd_ind].Rundennummer + '<br>';
                        if (runden_info[rd_ind].Name_Team === null) {
                            obs_text += runden_info[rd_ind].Startnr + ' - ' + runden_info[rd_ind].Dame + ' - ' + runden_info[rd_ind].Herr;
                        } else {
                            obs_text += runden_info[rd_ind].Startnr + ' - ' + runden_info[rd_ind].Name_Team;
                        }
                        if (runden_info[rd_ind].PpR === 2) {
                            obs_text += '<br>' + runden_info[rd_ind + 1].Startnr + ' - ' + runden_info[rd_ind + 1].Dame + ' - ' + runden_info[rd_ind + 1].Herr;
                        }
                        obs_text += '</div > <br><br><div class="wertung_offen" onclick="return senden(\'Runde_starten\',\'\');">Runde starten</div>';
                        HTML_erstellen.wait(rd_ind, runden_info, obs_text, wr_name, WR_ID, io);
                    }
                }
            }
            break;

        case "werten":		// Verteilung der verschiedenen Wertungsseiten
            var st_kl = runden_info[0].Startklasse;
            switch (wertungsrichter[WR_ID].WR_func) {
                case "X":
                    switch (st_kl.substring(0, 3)) {
                        case "BS_":
                            switch (st_kl) {
                                case "BS_BY_BJ":
                                case "BS_BY_BE":
                                case "BS_BY_BS":
                                case "BS_BY_S1":
                                    HTML_Seite = HTML_erstellen.BS_BY_BWSeite(rd_ind, runden_info, wr_name, WR_ID, wertungsrichter[WR_ID].WR_tausch, io);
                                    break;
                                case "BS_BW_BW":
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
                                    HTML_Seite = HTML_erstellen.BS_BW_BWSeite(rd_ind, runden_info, wr_name, WR_ID, wertungsrichter[WR_ID].WR_tausch, io);
                                    break;
                                default:
                                    HTML_Seite = HTML_erstellen.BS_Seite(rd_ind, runden_info, wr_name, WR_ID, wertungsrichter[WR_ID].WR_tausch, io);
                                    break;
                            }
                            break;
                        case "BW_":
                            if (new_guidelines) {
                                HTML_Seite = HTML_erstellen.BW_NG_Seite(rd_ind, runden_info, wr_name, WR_ID, wertungsrichter[WR_ID].WR_tausch, io);
                            } else {
                                HTML_Seite = HTML_erstellen.BW_Seite(rd_ind, runden_info, wr_name, WR_ID, wertungsrichter[WR_ID].WR_tausch, io);
                            }
                            setTimeout(function () {
                                io.sockets.emit('chat', WR_Info1);
                                io.sockets.emit('chat', WR_Info2);
                            }, 500);
                            break;
                        case "F_B":
                            HTML_Seite = HTML_erstellen.RR_Form_Seite(rd_ind, runden_info, akrobatiken, "X", wr_name, WR_ID, io);
                            break;
                    }
                    break;

                case "Ob":
                    st_kl = runden_info[0].Startklasse;
                    switch (st_kl.substring(0, 3)) {
                        case "BW_":
                        case "F_B":
                            HTML_Seite = HTML_erstellen.BW_Observer(rd_ind, runden_info, wr_name, WR_ID, io);
                            break;
                        case "RR_":
                        case "F_R":
                            if (runden_info[rd_ind].PpR === 1 && observer[WR_ID] > 0) {
                                HTML_erstellen.wait(rd_ind, runden_info, '<div class="wertung_offen" onclick="return p_logout()">kein Einsatz in dieser Runde</div>', wr_name, WR_ID, io);
                                io.sockets.emit('chat', { text: 'toRoot', WR: WR_ID });
                            } else {
                                HTML_erstellen.RR_Observer(rd_ind, runden_info, observer[WR_ID], wr_name, WR_ID, akrobatiken, anz_obs, io);
                            }
                            break;
                        case "BS_":
                            switch (st_kl) {
                                case "BS_BY_BJ":
                                case "BS_BY_BE":
                                case "BS_BY_BS":
                                case "BS_BY_S1":
                                    HTML_Seite = HTML_erstellen.BW_Observer(rd_ind, runden_info, wr_name, WR_ID, io);
                                    break;
                                case "BS_BW_BW":
                                case "BS_BW_SH":
                                    HTML_Seite = HTML_erstellen.BW_Observer(rd_ind, runden_info, wr_name, WR_ID, io);
                                    break;
                                default:
                                    HTML_Seite = HTML_erstellen.BW_Observer(rd_ind, runden_info, wr_name, WR_ID, io);
                                    break;
                            }
                            break;
                    }
                    HTML_beamer.beamer_runde(io, runden_info, runde);
                    HTML_moderator.runde(io, runden_info, runde);
                    break;

                case "Ft":
                    st_kl = runden_info[0].Startklasse;
                    switch (st_kl.substring(0, 3)) {
                        case "RR_":
                            if (st_kl === "RR_S1" || st_kl === " RR_S2") {
                                HTML_Seite = HTML_erstellen.MK_WB_Seite(rd_ind, runden_info, wr_name, WR_ID, wertungsrichter[WR_ID].WR_tausch, io, wertungsrichter[WR_ID].WR_func);
                           } else {
                                HTML_Seite = HTML_erstellen.RR_Seite(rd_ind, runden_info, akrobatiken, "Ft", wr_name, WR_ID, io);
                            }
                            break;
                        case "F_R":
                            HTML_Seite = HTML_erstellen.RR_Form_Seite(rd_ind, runden_info, akrobatiken, "Ft", wr_name, WR_ID, io);
                            break;
                        case "XXX":
                            break;
                    }
                    break;

                case "Ak":
                    st_kl = runden_info[0].Startklasse;
                    switch (st_kl.substring(0, 3)) {
                        case "RR_":
                            HTML_Seite = HTML_erstellen.RR_Seite(rd_ind, runden_info, akrobatiken, "Ak", wr_name, WR_ID, io);
                            break;
                        case "F_R":
                            HTML_Seite = HTML_erstellen.RR_Form_Seite(rd_ind, runden_info, akrobatiken, "Ak", wr_name, WR_ID, io);
                            break;
                        case "beamer_runde":
                            break;
                    }
                    break;

                case "MA":
                case "MB":
                    st_kl = runden_info[0].Startklasse;
                     HTML_Seite = HTML_erstellen.MK_WB_Seite(rd_ind, runden_info, wr_name, WR_ID, wertungsrichter[WR_ID].WR_tausch, io, wertungsrichter[WR_ID].WR_func);
                    break;

                default:
                    res.render(__dirname + '\\views\\pause.html');
                    break;
            }
            break;

        case "checked":
        case "wait":		// Warten zwischen den Tanzrunden WR
            HTML_erstellen.wait(rd_ind, runden_info, render_WR_status(), wr_name, WR_ID, io);
            break;

        case "checken":		// Grobfehler harmonisieren
            st_kl = runden_info[0].Startklasse;
            switch (st_kl.substring(0, 3)) {
                case "BW_":
                    HTML_Seite = HTML_erstellen.BW_ObsCheck(rd_ind, wertungsrichter, wertungen, runden_info, runde, wr_name, WR_ID, io, new_guidelines);
                    break;
                case "F_B":
                    HTML_Seite = HTML_erstellen.BW_ObsCheck(rd_ind, wertungsrichter, wertungen, runden_info, runde, wr_name, WR_ID, io, false);
                    break;
                case "F_R":
                case "RR_":
                    if (runden_info[rd_ind].PpR === 1 && observer[WR_ID] > 0) {
                        HTML_erstellen.wait(rd_ind, runden_info, '<div class="wertung_offen" onclick="return p_logout()">kein Einsatz in dieser Runde</div>', wr_name, WR_ID, io);
                    } else {
                        HTML_Seite = HTML_erstellen.RR_ObsCheck(rd_ind, wertungsrichter, wertungen, runden_info, runde, observer[WR_ID], wr_name, WR_ID, akrobatiken, anz_obs, io);
                    }
                    break;
                case "BS_":
                    auswerten("Obs_check1=Ok&WR_ID=" + WR_ID + "&rt_ID=" + runden_info[0].RT_ID);

                    break;
            }
            break;

        default:  // keine Wertung in dieser Runde
            HTML_erstellen.wait(rd_ind, runden_info, '<div class="wertung_offen" onclick="return p_logout()">kein Einsatz in dieser Runde</div>', wr_name, WR_ID, io);
            setTimeout(abmelden, 3000, WR_ID);
    }
    function abmelden(WR_ID) {
        io.sockets.emit('chat', { text: 'toRoot', WR: WR_ID });
    }
}

function auswerten(cgivar) {
    var temp = new Object();
    var wtext;
    var i;
    var rd_ind = 0;
    var wr_name;
    var body = cgi_split(cgivar);
    if (cgivar.indexOf("NaN") > 0) {
        wr_name = wertungsrichter[body.WR_ID].WR_Nachname.substr(0, 1) + wertungsrichter[body.WR_ID].WR_Vorname || "";
        console.log('Tablett von ' + wr_name + ' liefert falsche Werte!' + '\r\n' + '\r\n' + '\r\n');
    } else {
        if (body.Obs_check1 === "Ok") {
            for (i = 1; i < runde; i++) {
                rd_ind += parseInt(runden_info[i].PpR);
            }
            if (body.korr1 === "true" || body.korr2 === "true") {
                insert_differences(body, wertungen);
            }
            wtext = "";
            if (typeof body.TP_ID1 !== "undefined") {
                wtext += body.TP_ID1 + ';' + body.WR_ID + ';' + cgivar + '\r\n';
            }
            if (typeof body.TP_ID2 !== "undefined") {
                wtext += body.TP_ID2 + ';' + body.WR_ID + ';' + cgivar + '\r\n';
            }
            wtext = wtext.replace(/TP_ID/g, "PR_ID");
            //        wtext = wtext.replace(/-/g, "");
            fs.appendFileSync(conf.pfad + turnier_nr + body.rt_ID + '_raw.txt', wtext, encoding = 'utf8');
            fs.appendFileSync(conf.pfad + turnier_nr + body.rt_ID + '.txt', wtext, encoding = 'utf8');
            if (runden_info[rd_ind].PpR === 2) {
                wertungsrichter[body.WR_ID].WR_status = "checked";
                for (i in observer) {
                    if (wertungsrichter[i].WR_status !== "checked") {
                        verteilen(body.WR_ID);
                        refresh_wait();
                        return;
                    }
                }
            }
            HTML_auswerten.berechne_punkte(wertungen, runden_info, runde, wertungsrichter, conf.pfad + turnier_nr + body.rt_ID + '.txt');
            HTML_beamer.beamer_ranking(io, runden_info, runde);
            HTML_moderator.runde(io, runden_info, runde);

            runde++;
            WR_Info1 = new Object();
            WR_Info2 = new Object();
            for (i in wertungsrichter) {
                if (runden_info[0].Tanzrunde_MAX >= runde) {
                    if (wertungsrichter[i].WR_func === "Ob") {
                        wertungsrichter[i].WR_status = "runde";
                    }
                } else {
                    wertungsrichter[i].WR_status = "ende";
                }
                verteilen(i);
            }
            refresh_wait();
        } else {
            if (typeof body.TP_ID1 !== "undefined" || typeof body.TP_ID2 !== "undefined") {
                wertungsrichter[body.WR_ID].WR_status = "wait";
                if (typeof wertungen[body.WR_ID] === "undefined") {
                    temp = new Object();
                } else {
                    temp = wertungen[body.WR_ID];
                }
                var Punkte = 0;
                wtext = '';
                if (typeof body.TP_ID1 !== "undefined") {
                    Punkte = HTML_auswerten.rechne_wertungen(body, "1", runden_info);       // Punkte berechnen#
                    temp[body.TP_ID1] = { cgi: body, Punkte: Punkte, Runde: runde, Seite: 1 };

                    wtext += body.TP_ID1 + ';' + body.WR_ID + ';' + cgivar + '\r\n';
                }
                if (typeof body.TP_ID2 !== "undefined") {
                    Punkte = 0;
                    Punkte = HTML_auswerten.rechne_wertungen(body, "2", runden_info);       // Punkte berechnen
                    temp[body.TP_ID2] = { cgi: body, Punkte: Punkte, Runde: runde, Seite: 2 };

                    wtext += body.TP_ID2 + ';' + body.WR_ID + ';' + cgivar + '\r\n';
                }
                wtext = wtext.replace(/TP_ID/g, "PR_ID");
                fs.appendFileSync(conf.pfad + turnier_nr + body.rt_ID + '_raw.txt', wtext, encoding = 'utf8');
                //  WR Azubis nachfolgend nicht durchlaufen 
                wertungen[body.WR_ID] = temp;
                var count = 0;
                rundenende = false;
                for (var x in wertungen) {
                    for (i in wertungen[x]) {
                        if (parseInt(wertungen[x][i].cgi.rh1) === runde) {
                            count++;
                        }
                    }
                }
                for (i = 1; i < runde; i++) {
                    rd_ind += parseInt(runden_info[i].PpR);
                }
                if (Object.keys(observer).length === 2) {
                    if (anz_wr * runden_info[rd_ind].PpR === count + runden_info[rd_ind].PpR) {
                        rundenende = true;
                    }
                } else {
                    if (anz_wr * runden_info[rd_ind].PpR === count) {
                        rundenende = true;
                    }
                }

                if (rundenende === true) { 
                    if (runden_info[0].Tanzrunde_MAX >= runde) {
                        for (i in wertungsrichter) {
                            if (wertungsrichter[i].WR_func === "Ob") {
                                wertungsrichter[i].WR_status = "checken";
                                verteilen(i);
                            }
                        }
                    } else {
                        for (i in wertungsrichter) {
                            if (wertungsrichter[i].WR_func !== "") {
                                wertungsrichter[i].WR_status = "start";
                            }
                        }
                        WR_Info1 = new Object();
                        WR_Info2 = new Object();
                    }
                }
                //  WR Azubis ab hier weiterlaufen 
                refresh_wait();
            }
        }
    }
}

function refresh_wait() {
    render_WR_status();
    for (i in wertungsrichter) {
        if (wertungsrichter[i].WR_status === 'wait') {
            var wr_name = wertungsrichter[i].WR_Nachname.substr(0, 1) + wertungsrichter[i].WR_Vorname || "";
            HTML_erstellen.wait(0, runden_info, wr_status, wr_name, i, io);
        }
    }
    io.emit('chat', { msg: 'mod_wrstatus', text: wr_status });

}

function render_WR_status() {
    var content = "";
    var wr;
    for (var i in wertungsrichter) {
//        wr = wertungsrichter[i].WR_Vorname.substring(0, 1) + wertungsrichter[i].WR_Nachname.substring(0, 2);
        wr = wertungsrichter[i].WR_Kuerzel;
        switch (wertungsrichter[i].WR_status) {
            case "werten":
            case "runde":
               content = content + '<div class="wertung_offen">' + wr + '</div>';
                break;
            case "checken":
                content = content + '<div class="wertung_check">' + wr + '</div>';
                break;
            case "checked":
            case "wait":
                content = content + '<div class="wertung_ok">' + wr + '</div>';
                break;
            default:
        }
    }
//    io.emit('chat', { msg: 'aufWRwarten', text: content });
    io.emit('chat', { msg: 'beamer', wrstatus: content });
    wr_status = content;
    return content;
}

function to_Root() {
    setTimeout(function () {
        io.emit('chat', { msg: 'beamer', kopf:'', inhalt:'' });
        io.emit('chat', { msg: 'judgetool', text:  'toRoot' });
    }, 500);
}

function cgi_split(cgi) {
    var back = new Object();
    var sl = cgi.split("&");
    var teil;
    for (var t in sl) {
        teil = sl[t].split("=");
        back[teil[0]] = teil[1];
    }
    return back;
}

function turnier_titel() {
    connection
        .query('SELECT * FROM turnier;')
        .on('done', function (data) {
            title = data[0].Turnier_Name;
            turnier_nr = "T" + data[0].Turnier_Nummer + "_RT";
            wr_lesen();
        });
}

function wr_lesen() {
    connection
        .query('SELECT * FROM wert_richter ORDER BY WR_Kuerzel;')
        .on('done', function (data) {
            var wr = "";
            var anz = 0;
            anz_wr = 0;
            anz_obs = 0;
            var rd;
            var trunde;
            var st_kl;
            if (typeof runden_info[0] === "undefined") {
                st_kl = "RR_A";
                trunde = "VR";
                rd = "Vor_r";
            } else {
                st_kl = runden_info[0].Startklasse;
                trunde = runden_info[0].RundeArt;
                rd = runden_info[0].Runde;
            }
            var erster = false;
            observer = new Object;
            wertungsrichter = new Object();
            for (var i in data) {
                wertungsrichter[data[i].WR_ID] = data[i];
                if (wertungsrichter[data[i].WR_ID].WR_func === "Ob") {
                    if (trunde === 'ER' && erster === true && runden_info[0].Runde !== 'Semi') {
                        wertungsrichter[data[i].WR_ID].WR_func = "";
                        wertungsrichter[data[i].WR_ID].WR_status = "";
                    }
                    erster = true;
                }
                if (wertungsrichter[data[i].WR_ID].WR_func === "Ak" && (st_kl === "RR_S" || rd.indexOf("_Fu") > 0)) {
                    wertungsrichter[data[i].WR_ID].WR_func = "";
                    wertungsrichter[data[i].WR_ID].WR_status = "";
                }
                if (wertungsrichter[data[i].WR_ID].WR_func !== "") {
                    anz_wr++;
                }
                if (wertungsrichter[data[i].WR_ID].WR_func === "Ob") {
                    observer[data[i].WR_ID] = anz_obs;
                    anz_obs++;
                }
            }
            Index_Seite = HTML_erstellen.wr_login(wertungsrichter, title);
        });
}

turnier_titel();
console.log('App Started on PORT ' + conf.port + '.');
console.log('Version ' + ver);
to_Root();

function mk_verteilen() {
    connection
        .query('SELECT * FROM View_Rundenablauf WHERE Runde LIKE "MK_%" ORDER BY Startklasse, Startnr;')
//        .query('SELECT * FROM View_Rundenablauf WHERE Runde = "MK_3_BOT" ORDER BY Startklasse, Rundennummer, Startnr;')
        .on('done', function (data) {
            io.sockets.emit('chat', { msg: 'mk_fill', turnier: data[0].RT_file.substr(0, 8) });
            for (var i in data) {
                var mehrkampf = new Object;
                if (data[i].Runde.substring(0, 4) !== "MK_5") {
                    mehrkampf.RT_ID = data[i].RT_ID;
                    mehrkampf.TP_ID = data[i].TP_ID;
                    mehrkampf.Startklasse = data[i].Startklasse;
                    mehrkampf.Tanzrunde_Text = data[i].Tanzrunde_Text;
                    mehrkampf.Startnr = data[i].Startnr;
                    mehrkampf.Runde = data[i].Runde;
                    mehrkampf.Herr = data[i].Herr;
                    mehrkampf.Dame = data[i].Dame;
                    io.sockets.emit('chat', { msg: 'mk_fill', WR: '4', couple: mehrkampf });
                }
            }
            runden_info = data;		// Rundeninfo laden
        });
}