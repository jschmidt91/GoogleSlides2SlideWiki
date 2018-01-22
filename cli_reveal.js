"use strict";

/*
 * 2018-01-19:
 * 
 *
 * 2018-01-20:
 * TODO: Verlinkung von Inhalten übernehmen - erledigt
 * TODO: Verlinkung in Überschrift buggy - ggf. Stylesheet-Positionierung fehlerhaft.
 */

var fs = require('fs');
var request = require('request');
var htmlspecialchars = require('htmlspecialchars');

exports.converter = function(presentation) {
    var fileHTML = 'reveal.js/test.html';
    var fileJSON = 'presentationSlidewiki.json';

    var contentJSON = JSONcreateDeck(presentation);
    var contentHTML = HTMLcreateDeck(JSON.parse(contentJSON));


    writeFile(fileHTML, contentHTML);
    writeFile(fileJSON, contentJSON);
    writeFile('presentationGoogleAPI.json', JSON.stringify(presentation));
}

function writeFile(file, content) {
    fs.writeFile(file, content, function(err) {
        if (err) {
            return console.log(err);
        }

        //console.log("The file was saved!");
    });
}

function JSONcreateDeck(presentation) {
    var JSONdeck = {
        title: presentation.title,
        id: "95843",
        googleId: presentation.presentationId,
        type: "deck",
        user: "4327",
        theme: "undefined",
        children: []
    };

    if (typeof presentation.pageSize !== 'undefined') {
        JSONdeck.width = scaledEMU2PX(presentation.pageSize.width.magnitude, 1);
        JSONdeck.height = scaledEMU2PX(presentation.pageSize.height.magnitude, 1);
    }

    for (var i = 0; i < presentation.slides.length; i++) {
        JSONdeck.children.push(JSONcreateSlide(presentation.slides[i]));
    }

    return JSON.stringify(JSONdeck);
}

function JSONcreateSlide(slide) {
    var JSONslides = {
        title: "New slide",
        content: "",
        speakernotes: "",
        user: "4327",
        id: "1337",
        type: "slide",
        GoogleObjId: slide.objectId,
        slideElements: []
    };

    for (var i = 0; i < slide.pageElements.length; i++) {
        JSONslides.slideElements.push(JSONcreateSlideElement(slide.pageElements[i]));
        JSONslides.content += HTMLcreateSlideElement(JSONslides.slideElements[i]);
    }

    return JSONslides;
}

function JSONcreateSlideElement(element) {
    var JSONelement = {
        googleId: element.objectId,
        width: scaledEMU2PX(element.size.width.magnitude, element.transform.scaleX),
        height: scaledEMU2PX(element.size.height.magnitude, element.transform.scaleY)
    };

    if (typeof element.transform.translateX != 'undefined') {
        JSONelement.x = scaledEMU2PX(element.transform.translateX, 1)
    } else {
        JSONelement.x = 0;
    }

    if (typeof element.transform.translateY != 'undefined') {
        JSONelement.y = scaledEMU2PX(element.transform.translateY, 1)
    } else {
        JSONelement.y = 0;
    }

    if (typeof element.transform.shearX != 'undefined') {
        JSONelement.shearX = scaledEMU2PX(element.transform.shearX, 1);
    } else {
        JSONelement.shearX = 0;
    }

    if (typeof element.transform.shearY != 'undefined') {
        JSONelement.shearY = scaledEMU2PX(element.transform.shearY, 1);
    } else {
        JSONelement.shearY = 0;
    }


    if (typeof element.shape !== 'undefined' && typeof element.shape.placeholder !== 'undefined' && typeof element.shape.placeholder.type !== 'undefined') {
        JSONelement.contentType = element.shape.placeholder.type;
    } else if (typeof element.shape !== 'undefined' && typeof element.shape.shapeType !== 'undefined') {
        JSONelement.contentType = element.shape.shapeType;
    } else if (typeof element.table !== 'undefined') {
        JSONelement.contentType = 'TABLE';
    } else {
        JSONelement.contentType = 'NONE';
    }

    if (typeof element.shape !== 'undefined') {
        JSONelement.content = JSONcreateShape(element.shape, JSONelement.contentType);
    }

    if (typeof element.image !== 'undefined') {
        var imageTitle = '';

        if (typeof element.description != 'undefined') {
            imageTitle = element.description.replace(' ', '_');
        }

        JSONelement.image = JSONcreateImage(element.image, element.objectId, imageTitle);
        JSONelement.contentType = 'IMAGE';
    }

    if (typeof element.sheetsChart !== 'undefined') {
        var imageTitle = '';

        if (typeof element.title != 'undefined') {
            imageTitle = element.title.replace(' ', '_');
        }


        JSONelement.image = JSONcreateImage(element.sheetsChart, element.objectId, imageTitle);
        JSONelement.contentType = 'IMAGE';
    }

    if (typeof element.table !== 'undefined') {
        JSONelement.table = JSONcreateTable(element.table);
    }

    return JSONelement;
};

function scaledEMU2PX(emu, scale) {
    return Math.round(emu / 12700 * scale * 1.33333333333333);
};

function JSONcreateShape(shape, contentType) {
    var JSONcontent = {};

    if (shape.shapeType == 'TEXT_BOX' && typeof shape.text !== 'undefined' && typeof shape.text.textElements !== 'undefined') {
        JSONcontent = JSONcreateTextBox(shape.text, contentType);
    }

    return JSONcontent;
}

function JSONcreateTextBox(text, contentType) {
    var JSONcontent = [];
    var JSONtextElement = { "textType": "", "text": "", "html": "", "stylesheet": [] };

    var listId = '';
    var nestingLevel = 0;

    var paragraphMarker = '';

    for (var i = 0; i < text.textElements.length; i++) {
        // console.log(text.textElements[i]);
        if (typeof text.textElements[i].textRun != 'undefined') {}
        /*
         * Prüfen ob ParagraphMarker vorliegt
         * Indikator für ein neues Textelement z.B. Paragraph, Listen(-item), 
         * d.h. explizit nicht die Fortsetzung des letzten Textelements mit ggf. anderer Formatierung.
         */
        if (typeof text.textElements[i] !== 'undefined' && typeof text.textElements[i].paragraphMarker !== 'undefined') {
            /*
             * Prüfen ob ParagraphMarker ein ListenItem andeutet und speichert NestingLevel bzw. setzt es zurück,
             * sofern dem ListenItem kein NestingLevel zugeordnet werden kann.
             */
            if (typeof text.textElements[i].paragraphMarker.bullet !== 'undefined' && text.textElements[i].paragraphMarker.bullet.listId !== 'undefined') {
                listId = text.textElements[i].paragraphMarker.bullet.listId;
                if (typeof text.textElements[i].paragraphMarker.bullet.nestingLevel !== 'undefined') {
                    nestingLevel = text.textElements[i].paragraphMarker.bullet.nestingLevel;
                } else {
                    nestingLevel = 0;
                }

                paragraphMarker = 'listItem';
            } else {
                paragraphMarker = 'paragraph';
            }

            /*
             * Ein (neuer) ParagraphMarker indiziert ein neues Textelement, sodass das bisherige
             * abgeschlossen sein muss und gespeichert werden kann.
             *
             * Entsprechend wird das JSONtextElement in den leeren Anfangszustand zurückgesetzt.
             *
             * Eine Ausnahme stellt der erste Schleifendurchlauf dar, da vorher kein ParagraphMarker 
             * gesetzt sein kann.
             */
            if (i != 0) {
                JSONcontent.push(JSONtextElement);
                JSONtextElement = { "textType": "", "text": "", "html": "", "stylesheet": [] };
            }

            continue;
        }

        if (listId !== '') {
            JSONtextElement['textType'] = paragraphMarker;
            JSONtextElement['listId'] = listId;
            JSONtextElement['nestingLevel'] = nestingLevel;
            JSONtextElement['text'] += text.textElements[i].textRun.content.replace(/\n/, '').replace(/\u000b/, '\n');


            listId = '';
        } else if (typeof text.textElements[i].textRun !== 'undefined') {
            JSONtextElement['textType'] = paragraphMarker;
            JSONtextElement['text'] += text.textElements[i].textRun.content.replace(/\n/, '').replace(/\u000b/, '\n');


            if (typeof text.textElements[i].textRun.style !== 'undefined') {
                var JSONstylesheet = {};

                if (0 == JSONtextElement.stylesheet.length) {
                    JSONstylesheet.startChar = 0;
                } else {
                    JSONstylesheet.startChar = JSONtextElement.stylesheet[JSONtextElement.stylesheet.length - 1].endChar + 1;
                }

                JSONstylesheet.endChar = JSONtextElement.text.length;

                /*
                 * Wichtig für schließende Tags zur Formatierung
                 * vorher mit "continue;" bzgl. Doppelungen doch sinnvoll? Siehe unten gleiche Passage.
                 */

                if (JSONstylesheet.startChar > JSONstylesheet.endChar) {
                    JSONstylesheet.endChar = JSONstylesheet.startChar;
                }

                if (typeof text.textElements[i].textRun.style.bold !== 'undefined' && text.textElements[i].textRun.style.bold) {
                    JSONstylesheet.bold = true;
                } else {
                    JSONstylesheet.bold = false;
                }

                if (typeof text.textElements[i].textRun.style.italic !== 'undefined' && text.textElements[i].textRun.style.italic) {
                    JSONstylesheet.italic = true;
                } else {
                    JSONstylesheet.italic = false;
                }

                if (typeof text.textElements[i].textRun.style.fontSize !== 'undefined') {
                    JSONstylesheet.fontSize = text.textElements[i].textRun.style.fontSize.magnitude;
                    JSONstylesheet.fontSize += text.textElements[i].textRun.style.fontSize.unit.toLowerCase();
                } else {
                    JSONstylesheet.fontSize = false;
                }

                if (typeof text.textElements[i].textRun.style.link !== 'undefined') {
                    JSONstylesheet.link = text.textElements[i].textRun.style.link.url;
                }

                JSONtextElement.stylesheet.push(JSONstylesheet);
            }


            if ('TITLE' == contentType) {
                JSONtextElement['textType'] = 'heading3';
            }

        }

        /*
         * Stylesheets
         *
         * Prüfen, welche Textbereiche fettgedruckt oder kursiv geschrieben sind.
         */

        if (listId !== '' || typeof text.textElements[i].textRun !== 'undefined') {

            if (typeof text.textElements[i].textRun.style !== 'undefined') {
                var JSONstylesheet = {};
                if (0 == JSONtextElement.stylesheet.length) {
                    JSONstylesheet.startChar = 0;
                    JSONstylesheet.endChar = text.textElements[i].textRun.content.length - 1;
                } else {
                    JSONstylesheet.startChar = JSONtextElement.stylesheet[JSONtextElement.stylesheet.length - 1].endChar + 1;
                    JSONstylesheet.endChar = JSONtextElement.text.length - 1;
                }

                /*
                 * Wichtig für schließende Tags zur Formatierung
                 * vorher mit "continue;" bzgl. Doppelungen doch sinnvoll? Siehe oben gleiche Passage.
                 */
                if (JSONstylesheet.startChar > JSONstylesheet.endChar) {
                    JSONstylesheet.endChar = JSONstylesheet.startChar;
                }


                if (typeof text.textElements[i].textRun.style.bold !== 'undefined' && text.textElements[i].textRun.style.bold == true) {
                    JSONstylesheet.bold = true;
                } else {
                    JSONstylesheet.bold = false;
                }

                if (typeof text.textElements[i].textRun.style.italic !== 'undefined' && text.textElements[i].textRun.style.italic) {
                    JSONstylesheet.italic = true;
                } else {
                    JSONstylesheet.italic = false;
                }

                if (typeof text.textElements[i].textRun.style.link !== 'undefined') {
                    JSONstylesheet.link = text.textElements[i].textRun.style.link.url;
                }

                /*
                 * Bedingung bzgl. leerer Paragraphen, auf dem letzten Zeichen. TODO: Woher kommen die?
                 */
                if (JSONstylesheet.startChar != null) {
                    JSONtextElement.stylesheet.push(JSONstylesheet);
                }

            }
        }

        // //console.log(JSONtextElement.textType);
    }

    /*
     * Letztes JSONtextElement zum JSONcontent hinzufügen
     */
    JSONcontent.push(JSONtextElement);

    for (var k = 0; k < JSONcontent.length; k++) {
        /*
         * TextContent & Stylesheets auf HTML abbilden
         */
        var checkBold = false;
        var checkItalic = false;
        var checkLink = false;
        var checkLastLink = false;
        for (var j = 0; j < JSONcontent[k].stylesheet.length; j++) {
            if (JSONcontent[k].stylesheet[j].bold && checkBold == false) {
                checkBold = true;
                JSONcontent[k].html += '<b>';
            } else if (typeof JSONtextElement.stylesheet[j] != 'undefined' && JSONtextElement.stylesheet[j].bold == false && checkBold == true) {
                checkBold = false;
                JSONcontent[k].html += '</b>';
            }

            /*
             * Prüfen, ob zum Schluss der Tag geschlossen wird.
             */
            if (j + 1 == JSONcontent[k].stylesheet.length && true === checkBold) {
                JSONcontent[k].html += '</b>';
            }

            if (JSONcontent[k].stylesheet[j].italic && checkItalic == false) {
                checkItalic = true;
                JSONcontent[k].html += '<i>';
            } else if (JSONcontent[k].stylesheet[j].italic == false && checkItalic == true) {
                checkItalic = false;
                JSONcontent[k].html += '</i>';
            }

            /*
             * Prüfen, ob zum Schluss der Tag geschlossen wird.
             */
            if (j + 1 == JSONcontent[k].stylesheet.length && true === checkItalic) {
                JSONcontent[k].html += '</i>';
            }

            /*
             * Verlinkung prüfen
             */
            if(true === checkLink && (typeof JSONcontent[k].stylesheet[j].link == 'undefined' ||  (typeof JSONcontent[k].stylesheet[j].link != 'undefined' && JSONcontent[k].stylesheet[j].link != checkLastLink))){
            	checkLink = false;
            	checkLastLink = false;
            	JSONcontent[k].html += '</a>';
            }

            if(typeof JSONcontent[k].stylesheet[j].link != 'undefined' && JSONcontent[k].stylesheet[j].link != checkLastLink){
            	checkLink = true;
            	checkLastLink = JSONcontent[k].stylesheet[j].link;
            	JSONcontent[k].html += '<a href="'+JSONcontent[k].stylesheet[j].link+'">';
            }

            if('Website improvements - www.slidewiki.eu ' == JSONcontent[k].text){
            	console.log(JSONcontent[k].text[22]);
            	console.log(JSONcontent[k].text.substr(JSONcontent[k].stylesheet[j].startChar, JSONcontent[k].stylesheet[j].endChar - JSONcontent[k].stylesheet[j].startChar + 1));
            }

            JSONcontent[k].html += JSONcontent[k].text.substr(JSONcontent[k].stylesheet[j].startChar, JSONcontent[k].stylesheet[j].endChar - JSONcontent[k].stylesheet[j].startChar + 1);
        }
    }

    // Alte Variante ... Unzureichend, da nur letztes Element formatiert wurde
    // /*
    //      * TextContent & Stylesheets auf HTML abbilden
    //      */
    //     var checkBold = false;
    //     var checkItalic = false;
    //     for (var j = 0; j < JSONtextElement.stylesheet.length; j++) {
    //         if (JSONtextElement.stylesheet[j].bold && checkBold == false) {
    //             checkBold = true;
    //             JSONtextElement.html += '<b>';
    //         } else if (JSONtextElement.stylesheet[j].bold == false && checkBold == true) {
    //             checkBold = false;
    //             JSONtextElement.html += '</b>';
    //         }

    //         if (JSONtextElement.stylesheet[j].italic && checkItalic == false) {
    //             checkItalic = true;
    //             JSONtextElement.html += '<i>';
    //         } else if (JSONtextElement.stylesheet[j].italic == false && checkItalic == true) {
    //             checkItalic = false;
    //             JSONtextElement.html += '</i>';
    //         }
    //         JSONtextElement.html += JSONtextElement.text.substr(JSONtextElement.stylesheet[j].startChar, JSONtextElement.stylesheet[j].endChar - JSONtextElement.stylesheet[j].startChar + 1);
    //     }

    //console.log(JSONcontent);
    return JSONcontent;
}

function JSONcreateImage(image, objectId, imageTitle) {
    var JSONimage = {
        title: imageTitle,
        url: '../' + objectId + '_' + imageTitle
    };


    request(image.contentUrl).pipe(fs.createWriteStream(objectId + '_' + imageTitle)).on('close', function(err, res, body) {
        //console.log(res);
    });

    return JSONimage;
}

function JSONcreateTable(table) {
    var JSONtable = {
        rows: []
    };

    JSONtable.columns = table.columns;

    for (var i = 0; i < table.tableRows.length; i++) {
        JSONtable.rows.push(JSONcreateTableRow(table.tableRows[i]));
    }


    return JSONtable;
}

function JSONcreateTableRow(row) {
    var JSONtableRow = {
        cells: []
    };

    for (var i = 0; i < row.tableCells.length; i++) {
        JSONtableRow.cells.push(JSONcreateTableCell(row.tableCells[i]));
    }

    return JSONtableRow;
}

function JSONcreateTableCell(cell) {
    var JSONtableCell = {
        texts: []
    };

    for (var i = 0; i < cell.text.textElements.length; i++) {
        if (typeof cell.text.textElements[i].textRun !== 'undefined') {
            JSONtableCell.texts.push(JSONcreateTextElement('TABLE', undefined, cell.text.textElements[i]));
        }
    }

    return JSONtableCell;
}

function JSONcreateTextElement(shapeType, listType, textElement) {
    var JSONtextElement = {};

    var shapeText;

    if (typeof textElement.textRun !== 'undefined') {
        shapeText = htmlspecialchars(textElement.textRun.content);
        shapeText = shapeText.replace(/\n/, '').replace(/\u000b/, '\n');
    }

    if ('CENTERED_TITLE' == shapeType) {
        JSONtextElement.deckTitle = shapeText;
    }

    if ('SUBTITLE' == shapeType) {
        JSONtextElement.subtitle = shapeText;
    }

    if ('TITLE' == shapeType) {
        JSONtextElement.title = shapeText;
    }

    if ('BODY' == shapeType && listType === undefined) {
        JSONtextElement.paragraph = shapeText;
    }

    if ('BODY' == shapeType && listType == 'bulleted') {
        JSONtextElement.listElement = shapeText;
    }

    if ('NONE' == shapeType && listType === undefined) {
        JSONtextElement.paragraph = shapeText;
    }

    if ('TABLE' == shapeType) {
        JSONtextElement.text = htmlspecialchars(shapeText);
    }

    return JSONtextElement;
}

/*
 * Konstruktion: HTML - Content 
 * 
 * Nimmt die nach dem ETL-Prozess konstruierte SlideWikiJSON entgegen und erzeugt entsprechenden HTML-Code. 
 */

/*
 * Stand: 2017-12-11
 * HTML-Code wird derzeit aus GoogleAPIJSON erzeugt ... 
 * Ziel ist die Erzeugung aus der SlideWikiJSON.
 *
 * Stand: 2017-12-12
 * Anpassung an SlideWikiJSON-Notation, d.h. entsprechende Umbenennung von bspw.
 * "presentation" zu "deck" oder slides" zu "children", zur Orientierung am
 * Vokabular von SlideWiki zur internen Darstellung des HTML-Anzeige-Formats -
 * Erleichterung der Lesbarkeit von Projektteilnehmer.
 */

function HTMLcreateDeck(deck) {
    var content = '';

    content += HTMLcreateHeader(deck.title, deck.width, deck.height);

    for (var i = 0; i < deck.children.length; i++) {
        content += HTMLcreateSlide(deck.children[i], deck.width, deck.height);
    }

    content += HTMLcreateFooter();

    return content;
}

/*
 * Konstruktion: HTML-Header
 *
 * Erstellt zum gegebenen Titel des Decks das obere HTML-Grundgerüst bereit:
 * HTML-StartTag, Head-Tags
 * Title,
 * CSS,
 * Body - StartTag,
 * Reveal.js DIV-Container für die Klassen (u.a. CSS, JS) "reveal" und "slides"
 */

function HTMLcreateHeader(title, width, height) {
    return '<html>' +
        '<head>' +
        '<meta charset="utf-8">' +
        '<title>' + title + '</title>' +
        '<link rel="stylesheet" href="css/reveal.css">' +
        '<link rel="stylesheet" href="css/theme/white.css">' +
        '<link rel="stylesheet" href="../slideWiki.css">' +
        '</head>' +
        '<body>' +
        '<div class="reveal">' +
        '<div class="slides" style="width: ' + width + 'px; height: ' + height + 'px;">';

}

/*
 * 
 */
function HTMLcreateFooter() {
    return '</div>' +
        '</div>' +
        '<script src="lib/js/head.min.js"></script>' +
        '<script src="js/reveal.js"></script>' +
        '<script>' +
        'Reveal.initialize();' +
        '</script>' +
        '</body>' +
        '</html>'
}

function HTMLcreateSlide(slide, width, height) {
    var content = '';

    content += '<section style="width: ' + width + 'px; height: ' + height + 'px; font-size: 20pt; text-align: left;">';

    for (var i = 0; i < slide.slideElements.length; i++) {
        content += HTMLcreateSlideElement(slide.slideElements[i]);
    }

    content += '</section>';

    return content;
}

/*
 * Konstruktion: SlideElement
 *
 * Erstellt SlideElements entsprechend des vorliegenden ContentTypes,
 * d.h. generiert DIV-Container etc. für das Layout und 
 * verarbeitet ContentInhalt des Slides als wiederum einzelne Elemente.
 */
function HTMLcreateSlideElement(slideElement) {
    var content = '';

    /*
     * Stylesheets eigentl. vereinheitlichbar!
     * Prüfen, ob es an dieser Stelle ausreicht.
     *
     * stylesheet für verschiebbare Boxen
     * stylesheetLayout für Standardelemente
     */
    var stylesheet = 'position: absolute; ' +
        'width: ' + slideElement.width + 'px; ' +
        'height: ' + slideElement.height + 'px; ' +
        'top: ' + slideElement.y + 'px; ' +
        'left: ' + slideElement.x + 'px; ' +
        '';

    if (typeof slideElement.contentType !== 'undefined' && slideElement.contentType == 'TABLE') {
        var stylesheet = 'position: absolute; ' +
            'width: ' + slideElement.width * slideElement.table.columns + 'px; ' +
            'height: ' + slideElement.height + 'px; ' +
            'top: ' + slideElement.y + 'px; ' +
            'left: ' + slideElement.x + 'px; ' +
            'font-size: 16pt';
    }

    var stylesheetLayout = 'position: absolute; ' +
        'width: ' + slideElement.width + 'px; ' +
        'height: ' + slideElement.height + 'px; ' +
        'top: ' + slideElement.y + 'px; ' +
        'left: ' + slideElement.x + 'px; ' +
        '';

    if (typeof slideElement.contentType == 'undefined') {
        return '<h1>Fehler - Kein ContentType feststellbar!</h1>';
    }

    if (slideElement.contentType == 'CENTERED_TITLE') {
        content += '<div class="deck_title" style="' + stylesheet + '">';

        /*
         * TODO: deck_title bzgl. Schriftgröße aus Stylesheet holen
         */

        for (var i = 0; i < slideElement.content.length; i++) {
        	var fontSize = '';
        	if (typeof slideElement.content[i].stylesheet !== 'undefined' && slideElement.content[i].stylesheet.length > 0 && typeof slideElement.content[i].stylesheet[0].fontSize != 'undefined' && false !== typeof slideElement.content[i].stylesheet[0].fontSize) {
		        fontSize = 'font-size: ' + slideElement.content[i].stylesheet[0].fontSize + '; ';
		        if(false == slideElement.content[i].stylesheet[0].fontSize){
		        	fontSize = 'font-size: 32pt; ';
		        }
		    }


            content += '<h1 style="' + fontSize + '">' + slideElement.content[i].text + '</h1>';
        }

        content += '</div>';
    }

    if (slideElement.contentType == 'SUBTITLE') {
        content += '<div class="deck_subtitle" style="' + stylesheet + '">';

        for (var i = 0; i < slideElement.content.length; i++) {
        	var fontSize = '';
        	if (typeof slideElement.content[i].stylesheet !== 'undefined' && slideElement.content[i].stylesheet.length > 0 && typeof slideElement.content[i].stylesheet[0].fontSize != 'undefined' && false !== typeof slideElement.content[i].stylesheet[0].fontSize) {
		        fontSize = 'font-size: ' + slideElement.content[i].stylesheet[0].fontSize + '; ';
		    }
            content += '<h' + (i + 2) + ' style="'+fontSize+'">' + slideElement.content[i].text + '</h' + (i + 2) + '>';
        }

        content += '</div>';
    }

    if (slideElement.contentType == 'TITLE') {
        content += '<div class="slide_title" style="' + stylesheetLayout + '">';

        for (var i = 0; i < slideElement.content.length; i++) {
            if (typeof slideElement.content[i].stylesheet[0].fontSize !== 'undefined' && slideElement.content[i].stylesheet[0].fontSize !== false) {
                content += '<h' + (i + 3) + ' style="font-size: ' + slideElement.content[i].stylesheet[0].fontSize + ';">' + slideElement.content[i].text + '</h' + (i + 3) + '>';
            } else {
                content += '<h' + (i + 3) + ' style="font-size: ' + (28 - 2 * i) + 'pt;">' + slideElement.content[i].text + '</h' + (i + 3) + '>';
            }
        }

        content += '</div>';
    }

    if (slideElement.contentType == 'BODY' || slideElement.contentType == 'TEXT_BOX') {
        if ('BODY' == slideElement.contentType) {
            var stylesheet = 'position: absolute; width: ' + slideElement.width + 'px; ' +
                'height: ' + slideElement.height + 'px; ' +
                'top: ' + slideElement.y + 'px; ' +
                'left: ' + slideElement.x + 'px; ' +
                '';

            content += '<div class="text_box" style="' + stylesheet + '">';
        } else if ('TEXT_BOX' == slideElement.contentType) {
            var stylesheet = 'position: absolute; ' +
                'width: ' + slideElement.width + 'px; ' +
                'height: ' + slideElement.height + 'px; ' +
                'top: ' + slideElement.y + 'px; ' +
                'left: ' + slideElement.x + 'px; ' +
                '';

            content += '<div class="text_box" style="' + stylesheet + '">';
        }

        var listId = '';
        var listCounter = 0;
        var nestingLevel = 0;

        for (var i = 0; i < slideElement.content.length; i++) {
            /*
             * Prüfen, ob das aktuelle TextElement zu einer Liste gehört.
             */
            if (slideElement.content[i].textType == 'listItem') {


                /*
                 * Prüfen, ob das aktuelle ListenTextElement zu einer neuen Liste gehört.
                 */
                if (listId != slideElement.content[i].listId) {
                    listId = slideElement.content[i].listId;
                    content += '<ul>';
                    listCounter = 0;
                }

                if (slideElement.content[i].nestingLevel > nestingLevel) {
                    content += '<ul>';
                    nestingLevel = slideElement.content[i].nestingLevel;
                    listCounter = 0;
                } else if (slideElement.content[i].nestingLevel < nestingLevel) {
                    for(var k=0; k<(nestingLevel-slideElement.content[i].nestingLevel);k++){
                    	content += '</li></ul></li>';
                    }
                    nestingLevel = slideElement.content[i].nestingLevel;
                    listCounter = 0;
                }

                if (listCounter != 0) {
                    content += '</li>';
                }
                listCounter++;

                content += '<li>' + slideElement.content[i].html + '';
            }

            /*
             * Wenn ein Wechsel von einem Liste-TextElement zu einem NichtListen-TextElement wechselt,
             * so muss die vorhergehende Liste geschlossen werden.
             */
            if (slideElement.content[i].textType != 'listItem' && listId != '') {
                for (var j = 0; j < nestingLevel; j++) {
                    content += '</li></ul>';
                }
                nestingLevel = 0;

                content += '</li></ul>';
                listId = '';
            }

            if (slideElement.content[i].textType == 'paragraph') {
                var stylesheetParagraph = '';

                if (typeof slideElement.content[i].stylesheet !== 'undefined' && slideElement.content[i].stylesheet.length > 0 && typeof slideElement.content[i].stylesheet[0].fontSize !== 'undefined' && slideElement.content[i].stylesheet[0].fontSize !== false) {
                    var stylesheetParagraph = 'font-size: ' + slideElement.content[i].stylesheet[0].fontSize + ';';
                }

                content += '<p style="' + stylesheetParagraph + '">' + slideElement.content[i].html + '</p>';
            }

        }

        /*
         * Falls letztes Textelement im ContentElement zu einer Liste gehört, 
         * so muss diese abgeschlossen werden.
         * 
         * Gleichermaßen müssen geschachtelte Listen vollständig, wohlgeformt geschlossen werden.
         */
        for (var j = 0; j < nestingLevel; j++) {
            content += '</li></ul>';
        }

        if (listId != '') {
            content += '</li></ul>';
        }

        content += '</div>'
    }

    if (slideElement.contentType == 'TABLE') {
        //console.log(slideElement.table);
        content += '<table style="' + stylesheet + '">';
        for (var k = 0; k < slideElement.table.rows.length; k++) {
            content += '<tr>';

            for (var l = 0; l < slideElement.table.rows[k].cells.length; l++) {
                for (var m = 0; m < slideElement.table.rows[k].cells[l].texts.length; m++) {
                    content += '<td>' + slideElement.table.rows[k].cells[l].texts[m].text + '</td>';
                }
            }

            content += '</tr>';
        }
        content += '</table>'
    }

    if (slideElement.contentType == 'IMAGE') {
        //console.log('\n\nTest: '+pageElement.description.replace(' ',''));
        var imageTitle = slideElement.image.title;

        //if(typeof pageElement.description != 'undefined') {
        //    imageTitle = imageTitle+'_'+pageElement.description.replace(' ','');
        //}

        //content += HTMLcreateImage(pageElement.image, imageTitle);

        /*
         * Positionierung
         */
        var stylesheet = 'position: absolute; ' +
            'margin-top: 0px; ' +
            'margin-right: 0px; ' +
            'margin-bottom: 0px; ' +
            'margin-left: 0px; ' +
            'border: 0px;';

        var stylesheetBox = 'position: absolute; ' +
            'overflow: hidden;' +
            'width: ' + slideElement.width + 'px; ' +
            'height: ' + slideElement.height + 'px; ' +
            'top: ' + slideElement.y + 'px; ' +
            'left: ' + slideElement.x + 'px; ';

        content += '<div style="' + stylesheetBox + '"><img src="' + slideElement.image.url + '" title="' + slideElement.image.title + '" alt="' + slideElement.image.title + '"  style="' + stylesheet + '"></div>';
    }

    /*
    if (typeof pageElement.shape !== 'undefined') {
        content += HTMLcreateShape(pageElement.shape);
    }

    if (typeof pageElement.image !== 'undefined') {
        //console.log('\n\nTest: '+pageElement.description.replace(' ',''));
        var imageTitle = pageElement.objectId;

        //if(typeof pageElement.description != 'undefined') {
        //    imageTitle = imageTitle+'_'+pageElement.description.replace(' ','');
        //}

        content += HTMLcreateImage(pageElement.image, imageTitle);
    }

    if (typeof pageElement.sheetsChart !== 'undefined') {
        //console.log('\n\nTest: '+pageElement.title.replace(' ',''));
        var imageTitle = pageElement.objectId;

        
        //if(typeof pageElement.description != 'undefined') {
        //    imageTitle = imageTitle+'_'+pageElement.title.replace(' ','');
        //}

        content += HTMLcreateImage(pageElement.sheetsChart, imageTitle);
    }

    if (typeof pageElement.table !== 'undefined') {
        content += HTMLcreateTable(pageElement.table);
    }
    */




    return content;
};

function HTMLcreateShape(shape) {
    var content = '';

    var shapeType;
    if (typeof shape.placeholder !== 'undefined') {
        if (typeof shape.placeholder.type !== 'undefined') {
            shapeType = shape.placeholder.type;
        }
    } else if (typeof shape.table != 'undefined') {
        shapeType = 'TABLE';

    } else {
        shapeType = 'NONE';
    }

    if (typeof shape.text !== 'undefined') {
        var listType = undefined;
        var nestingLevel = 0;
        var nestingLevelNew = 0;

        if (typeof shape.text.lists !== 'undefined') {
            listType = 'bulleted';
            content += '<ul>';
        }

        for (var i = 0; i < shape.text.textElements.length; i++) {
            if (listType == 'bulleted') {
                if (typeof shape.text.textElements[i].paragraphMarker !== 'undefined') {
                    if (typeof shape.text.textElements[i].paragraphMarker.bullet !== 'undefined') {
                        if (typeof shape.text.textElements[i].paragraphMarker.bullet.nestingLevel !== 'undefined') {
                            nestingLevelNew = shape.text.textElements[i].paragraphMarker.bullet.nestingLevel;
                        } else {
                            nestingLevelNew = 0;
                        }
                    }
                }
            }

            if (typeof shape.text.textElements[i].paragraphMarker !== 'undefined' && nestingLevel < nestingLevelNew) {
                content = content.substr(0, content.lastIndexOf('</li>'));
                content += '<ul>';
            }
            content += HTMLcreateTextElement(shapeType, listType, shape.text.textElements[i]);

            if (typeof shape.text.textElements[i].paragraphMarker !== 'undefined' && nestingLevel > nestingLevelNew) {
                content += '</ul></li>';
            }
            nestingLevel = nestingLevelNew;
        }

        for (var i = 0; i < nestingLevel; i++) {
            content += '</ul></li>';
        }

        if (listType == 'bulleted') {
            content += '</ul>';
        }
    }

    return content;
}

function HTMLcreateTextElement(shapeType, listType, textElement) {
    var content = '';

    var shapeText;

    if (typeof textElement.textRun !== 'undefined') {
        shapeText = htmlspecialchars(textElement.textRun.content);
        shapeText = shapeText.replace(/\n/, '');
    } else {
        return '';
    }

    /*
     * Dopplungen!? Bzgl. Headlines
     */
    if ('CENTERED_TITLE' == shapeType) {
        content += '<h1 style="font-size: 32pt;">' + shapeText + '</h1>';
    }

    if ('SUBTITLE' == shapeType) {
        content += '<h2 style="font-size: 30pt;">' + shapeText + '</h2>';
    }

    if ('TITLE' == shapeType) {
        content += '<h3 style="font-size: 28pt;">' + shapeText + '</h3>';
    }

    if ('BODY' == shapeType && listType === undefined) {
        content += '<p>' + shapeText + '</p>';
    }

    if ('BODY' == shapeType && listType == 'bulleted') {
        content += '<li>' + shapeText + '</li>';
    }

    if ('NONE' == shapeType && listType === undefined) {
        content += '<p>' + shapeText + '</p>';
    }

    if ('TABLE' == shapeType) {
        content += shapeText
    }

    return content;
}

function HTMLcreateImage(image, filename) {
    var self = this;

    var content = '';
    var fileType = '';

    request = require('request');

    /*
    var download = function(uri, filename, callback){
        fileType = 'Testtype';
        var fileType2 = '';
        var test = request.head(uri, function(err, res, body){
            //console.log('content-length:', res.headers['content-length']);
            if('image/png' == res.headers['content-type']){
                fileType2 = '.png';
            }

            if('image/jpeg' == res.headers['content-type']){
                fileType2 = '.jpg';
            }

            //request(uri).pipe(fs.createWriteStream(filename+fileType2)).on('close', function(){});
        });


        //console.log('Test: '+fileType2);
    };
    
    download(image.contentUrl, fileName, function(){
      //console.log('Image saved!');
    });*/


    request(image.contentUrl).pipe(fs.createWriteStream(filename)).on('close', function(err, res, body) {
        //console.log(res);
    });

    //console.log(request.head(image.contentUrl));
    //console.log("FileType::"+fileType);
    content = '<div><img style="border: 0px;" src="../' + filename + '"></div>';

    return content;
}

function HTMLcreateTable(table) {
    var content = '';

    content += '<table border="1" cellpadding="1" cellspacing="1">';
    for (var i = 0; i < table.tableRows.length; i++) {
        content += HTMLcreateTableRow(table.tableRows[i]);
    }

    content += '</table>';

    return content;
}

function HTMLcreateTableRow(row) {
    var content = '';

    content += '<tr>';

    for (var i = 0; i < row.tableCells.length; i++) {
        content += HTMLcreateTableCell(row.tableCells[i]);
    }

    content += '</tr>';

    return content;
}

function HTMLcreateTableCell(cell) {
    var content = '';

    content += '<td>';

    for (var i = 0; i < cell.text.textElements.length; i++) {

        content += HTMLcreateTextElement('TABLE', undefined, cell.text.textElements[i]);
    }

    content += '</td>';

    return content;
}