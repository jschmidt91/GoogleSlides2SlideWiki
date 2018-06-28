"use strict";

var fs = require("fs");
var JSZip = require("jszip");

// read a zip file
fs.readFile("reveal.zip", function(err, data) {
    JSZip.loadAsync(data).then(function(zip) {
        Object.keys(zip.files).forEach(function(filename) {
            if (filename == "reveal.js/test.html") {
                zip.files[filename].async('string').then(function(fileData) {
                    Reveal2JSON(fileData);
                })
            }
        })
    })
});


//Reveal2JSON
function Reveal2JSON(html) {
    // Objective: Output-Format like deckservice.{environment}.slidewiki.org/deck/{id}/slides 
    var JSONdeck = {};

    // Objective: Increment of count of slides in the deck for JSON.deck.slidesCount
    var slideCounter = 0;

    // Objective: auxiliary variable for string positions w.r.t. to start, end, area (source code between start and end) of e.g. sections
    var start, end, area;

    // Objective: Check existence of html head section, e.g. for the possibility of extracting title tags, themes, etc.
    /*
     * Extracts source code between of the head section <html>{source_code}</html>
     */
    var htmlHead = '';
    if (html.includes('<head>') && html.includes('</html>')) {
        htmlHead = html.substr(html.indexOf('<head>') + 6, html.indexOf('</head>') - html.indexOf('<head>') - 6);
    }
    //console.log(htmlHead);

    // Objective: extract title from html head section
    /*
     * Default-value: 'title tag missing / incorrect' if head section or title tag doesn't exists.
     * 
     * Extracts title <title>{title}</title>
     */
    JSONdeck.title = 'title tag missing / incorrect';
    if (htmlHead.includes('<title>') && htmlHead.includes('</title>')) {
        JSONdeck.title = htmlHead.substr(htmlHead.indexOf('<title>') + 7, htmlHead.indexOf('</title>') - htmlHead.indexOf('<title>') - 7);
    }
    //console.log(JSONdeck.title);

    // Objective: GET deck id from System later - TODO
    JSONdeck.id = 42;

    // Objective: default value, cf. deckservice.{environment}.slidewiki.org/deck/{id}/slides
    JSONdeck.type = 'deck';

    // Objective: GET user id from System later - TODO
    JSONdeck.user = '1337';

    // Objective: Extract reveal theme
    /*
     * default value 'default'
     * 
     * check link tags in html head to extract theme
     */
    JSONdeck.theme = 'default';
    // Objective: indexOf for the next occurrence
    var startIndex = 0;
    // Descriminator: link href of stylesheets defines the reveal css theme
    var href;
    while (-1 < htmlHead.indexOf('<link', startIndex)) {
        start = htmlHead.indexOf('<link', startIndex);
        end = htmlHead.indexOf('>', start);
        area = htmlHead.substr(start+5, end - start - 5).trim();
        //console.log(area);

        // Objetice: Check if the current link tag contains stylesheets and an attribute for an url
        if (area.includes('rel="stylesheet"') && area.includes('href="')) {
            // Obective: Extracts url for stylesheet file
            start = area.indexOf('href="');
            end = area.indexOf('"', start+6);
            href = area.substr(start+6, end - start - 6);

            // theme assignment
            /*
             * Check & Compate current URL with standard themes of reveal
             *
             * (!) reveal default reveal theme is black vs. slide wiki white
             * TODO: check values for themes in the slide wiki plattform
             */
            switch(href) {
                case 'css/theme/black.css':
                    JSONdeck.theme = 'black';
                    break;
                case 'css/theme/white.css':
                    JSONdeck.theme = 'default';
                    break;
                default:
                    JSONdeck.theme = 'default';
            }

            //console.log(start, end, href);
        }
        //console.log(JSONdeck.theme);

        // Objective: update startIndex to find the next link tag in the following iteration
        startIndex = htmlHead.indexOf('<link', startIndex) + 1;
    }

    // Objective: childen array, cf. deckservice.{environment}.slidewiki.org/deck/{id}/slides
    JSONdeck.children = [];

    // Objective: Betrachten von 
    var html = html.substr(html.indexOf('<section'), html.length - html.indexOf('<section'));

    // Objective: each section area represents an element of children = [] 
    while (html.indexOf('<section') > -1) {
        // var for the children element
        var JSONslide = {};

        // Objective: default value
        /*
         * Idea: check H1, H2, etc. tag for slide title informations
         */
        JSONslide.title = 'New slide';

        // Position of start and end of the current section
        start = html.indexOf('<section'); // currenlty, not in use
        end = html.indexOf('</section>');

        // Objective: extract HTML content for children element
        /*
         * open section tag ends with ">", i.e. we needs the html source code between this element and the section end tag position
         */
        JSONslide.content = html.substr(html.indexOf('>') + 1, end - html.indexOf('>') - 1)

        // TODO: extract spreakernotes
        /*
         * Option 1 - extract from attribute: <section data-notes="Something important">
         * Option 2 - extract from <aside class="notes">{speaker_notes}</aside) within the sections
         */
        JSONslide.speakernotes = '';

        // Objective: Until the creation, the user of the decks / slides would bethe same
        JSONslide.user = JSONdeck.user;

        // Objective: slide id w.r.t. deck id & additional counter / increment
        JSONslide.id = JSONdeck.id + "-" + slideCounter;

        // Objective: reveal defines an overall theme in the html head, so it should be the same here
        JSONslide.theme = JSONdeck.theme;

        // Objective: default value for slides
        JSONslide.type = "slide";

        // Objective: remove the processed html parts / sections
        html = html.substr(end + 10, html.length - end + 10);

        // Objective: push the final slide element to children[] 
        JSONdeck.children.push(JSONslide);

        // Objective: Increment the slide counter for the overall slidesCount of the deck and next slide id
        slideCounter++;
    }

    JSONdeck.slidesCount = slideCounter;

    //console.log(JSONdeck);

    // Objective: Export JSON
    var fileJSON = 'presentationSlidewikiRevealImport.json';
    writeFile(fileJSON, JSON.stringify(JSONdeck));

}


function writeFile(file, content) {
    fs.writeFile(file, content, function(err) {
        if (err) {
            return console.log(err);
        }

        //console.log("The file was saved!");
    });
}