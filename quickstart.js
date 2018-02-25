/**
 * Require.js
 * Datei- und Modul-Lader
 */
var cli_reveal = require('./cli_reveal');

/**
 * Node.js File System Module
 * Arbeit mit Dateisystem auf Server/Computer
 * Read, Create, Update, Delete, Rename files
 */
var fs = require('fs');

/**
 * Node.js Readline Module
 * Zeilenweises Lesen eines Datenstroms
 */
var readline = require('readline');

/**
 * Google API Service Client
 */
var google = require('googleapis');
var googleAuth = require('google-auth-library');

/**
 * OAuth 2.0 Scopes for Google APIs
 * Anzeigen der eigenen Google Slides Präsentationen
 * https://www.googleapis.com/auth/presentations.readonly
 * 
 * Vollständige Verwaltung der Google Slides Präsentationen
 * https://www.googleapis.com/auth/presentations
 *
 * Anzeigen der eigenen Google Drive Daten
 * https://www.googleapis.com/auth/drive.readonly
 * 
 * Vollständige Verwaltung der Google Drive Daten
 * https://www.googleapis.com/auth/drive
 */
var SCOPES = ['https://www.googleapis.com/auth/presentations.readonly'];

/**
 * TOKEN zur Identifizierung und Authentifizierung
 * process.env.HOME, process.env.HOMEPATH, process.env.USERPROFILE
 * Enthält Dateipfad zum Stammverzeichnis des Dateisystems auf Server/Computer
 * 
 * TOKEN_DIR Verzeichnis mit Authentifizierungsdatei
 *
 * TOKEN_PATH Pfad zur Authentifizierungsdatei im Format ".json"
 * Enthält access_token, refresh_token, token_type, expiry_date (Ablaufdatum)
 */
var TOKEN_DIR = (process.env.HOME || process.env.HOMEPATH ||
    process.env.USERPROFILE) + '/.credentials/';
var TOKEN_PATH = TOKEN_DIR + 'slides.googleapis.com-nodejs-quickstart.json';

/**
 * Laden der Datei: "client_secret.json"
 * installed: client_id, project_id, auth_uri, token_uri, auth_provider_x509_cert_url, client_secret, redirect_uris, 
 *
 */
 // fs.readFile(path[, options], callback)
 // callback <Function>: err <Error> data <string> | <Buffer>
fs.readFile('client_secret.json', function processClientSecrets(err, content) {
  if (err) {
    console.log('Fehler beim Laden der client_secret.json: ' + err);
    return;
  }
  // Authorisierung des Clients mit den gegebenen credentials: 
  // Danach Anfrage (call) der Google Slides API.
  // JSON.parse wandelt JSON-Datei in JavaScript-Objekt
  authorize(JSON.parse(content), listSlides);
});

/**
 * Erstellen eines OAuth2 Clients mit den gegebenen credentials:
 * Anschließend Ausführen der callback-Funktion
 *
 * @param {Object} credentials The authorization client credentials (parsed client_secret.json)
 * @param {function} callback The callback to call with the authorized client (function listSlides)
 */
function authorize(credentials, callback) {
  // Inhalte aus client_secret.json
  var clientSecret = credentials.installed.client_secret;
  var clientId = credentials.installed.client_id;
  var redirectUrl = credentials.installed.redirect_uris[0];

  // Authentifizierung OAuth2 Client erstellen
  var auth = new googleAuth();
  var oauth2Client = new auth.OAuth2(clientId, clientSecret, redirectUrl);

  // REST Resources
  var REST;

  // Prüfen ob bereits ein gespeicherter TOKEN existiert
  fs.readFile(TOKEN_PATH, function(err, token) {
    if (err) {
      getNewToken(oauth2Client, callback);
    } else {
      // Übergabe des credentials aus TOKEN
      oauth2Client.credentials = JSON.parse(token);
      // Aufruf listSlides()
      callback(oauth2Client);
      REST = 'Test';
    }
  });
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 *
 * @param {google.auth.OAuth2} oauth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback to call with the authorized
 *     client.
 */
function getNewToken(oauth2Client, callback) {
  var authUrl = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES
  });
  console.log('Authorize this app by visiting this url: ', authUrl);
  var rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });
  rl.question('Enter the code from that page here: ', function(code) {
    rl.close();
    oauth2Client.getToken(code, function(err, token) {
      if (err) {
        console.log('Error while trying to retrieve access token', err);
        return;
      }
      oauth2Client.credentials = token;
      storeToken(token);
      callback(oauth2Client);
    });
  });
}

/**
 * Store token to disk be used in later program executions.
 *
 * @param {Object} token The token to store to disk.
 */
function storeToken(token) {
  try {
    fs.mkdirSync(TOKEN_DIR);
  } catch (err) {
    if (err.code != 'EEXIST') {
      throw err;
    }
  }
  fs.writeFile(TOKEN_PATH, JSON.stringify(token));
  console.log('Token stored to ' + TOKEN_PATH);
}

/**
 * Anfragen der REST Resources zur Google Slides Präsentation mit presentationID
 */
function listSlides(auth) {
  // Spezifikation Google Slides API Call
  var slides = google.slides('v1');


  slides.presentations.get({
    auth: auth,
    //presentationId: '197DB-DQj0_cJEX4PfXl93xnutDd7TocnnoSY9UTIS3A'
    //presentationId: '13onqxdVYqGU0UU2ffLFWliw5Yg7gkjpJ42rZl5aBBNg'
    presentationId: '1PXxdBTWHmI_1fYrxJw0_b1b1GR92lrRLwagqxkkqysk'
    // presentationId: '1YDbb_5NElac1Xi8_3BSifEGwnLd6NZGBi0b4QxwTgEs'
  }, function(err, presentation) {
    if (err) {
      console.log('The API returned an error: ' + err);
      return;
    }

/*
    // Anzahl an Folien und enthaltenen Elemente
    
    var length = presentation.slides.length;
    console.log('The presentation contains %s slides:', length);
    for (i = 0; i < length; i++) {
      var slide = presentation.slides[i];
      console.log('- Slide #%s contains %s elements.', (i + 1),
          slide.pageElements.length)
    }
    */
    cli_reveal.converter(presentation);
    //console.log(presentation);
  });


}
