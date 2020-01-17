/**
 * @preserve Code.gs
 * - part of Classroom Activity Utility (CAU) (Google Doc add-on)
 * - Copyright (c) 2014-2015 Clif Kussmaul, clif@kussmaul.org, clif@cspogil.org, clifkussmaul@gmail.com
 */

// NOTE: 90%+ of execution time is used by API calls to Document, DocumentApp, File, & Folder
//       when CAU tries to change permissions, copy, open, remove

// URGENT: trivial example document with instructions - add to wiki page, use as screenshot
// URGENT: stressful test document - same file as add-on code
// URGENT: screenshot (simple example document), promotional image

// TODO: set output file for student and/or teacher version
//       https://ctrlq.org/code/20039-google-picker-with-apps-script
// TODO: generate set of N student activities (e.g. with team names)
//       prompt or look in suffix for ;-list of values - ? different sets for different courses 

// TODO: add tag to use in header of teacher versions
// TODO: auto row labels (letters) in question tables
// TODO: auto line numbers in code listings (from current cursor)
//       get current position with Document.getCursor(), find max lines in any column, generate numbers

// TODO: allow regex for AUTHOR_SUFFIX to allow (Master), (Author), etc.
// TODO: measure execution time and display at end (relative to Google 6 minute limit)
// TODO: respond to Google feedback
// TODO: recheck style & CSS guides: https://developers.google.com/apps-script/add-ons/{style,css}
// TODO: recheck JSDoc tags:         https://developers.google.com/closure/compiler/docs/js-for-compiler

// TODO: finish {{INSERT filename}} at specified positions (e.g. names & roles; reflector report; rubric)
//       - insert different content for students (e.g. header & roles) & teachers (e.g. feedback request)
// TODO: generate checkup report in separate document, not just sidebar 
// TODO: better error handling (messages & recovery) when unable to remove old files or change permissions 
// TODO: menu option to prompt for set of versions to make
// TODO: ? should properties be per-doc (better for shared docs) or per-user (easier to update)

// from: http://tinyurl.com/copydocscript
// function copyDocs() {
//  for(i=0; i<60; i++){  
//      DriveApp.getFileById('1ZNYl6FwS5LWbYfLGAGG8IBZEY89aQeoMCwhrWczWckE').makeCopy();
//    }
//  }

// FUTURE: run on all or selected docs in a folder (& subfolders?)

// FUTURE: edit in Eclipse - if and when Eclipse supports scripts in docs, rather than only standalone scripts
// FUTURE: when supported by Google Apps Script
// - set headings, especially ClearHeading
// - set other formatting (e.g. tables spacing, coloring, line styles)
// - set tags USER, EMAIL, PAGE


//********************************************************************************
// global variables

// list of paragraph heading names and enumeration values - order matters!
// (can't find a way to iterate through JavaScript enum values)
var HEADINGS = { 
  'Title'    : DocumentApp.ParagraphHeading.TITLE,
  'Subtitle' : DocumentApp.ParagraphHeading.SUBTITLE,
  'Heading1' : DocumentApp.ParagraphHeading.HEADING1,
  'Heading2' : DocumentApp.ParagraphHeading.HEADING2,
  'Heading3' : DocumentApp.ParagraphHeading.HEADING3,
  'Heading4' : DocumentApp.ParagraphHeading.HEADING4,
  'Heading5' : DocumentApp.ParagraphHeading.HEADING5,
  'Heading6' : DocumentApp.ParagraphHeading.HEADING6,
  'Normal'   : DocumentApp.ParagraphHeading.NORMAL,
};

//********************************************************************************
// top-level callback functions

/**
 * When script is installed, open. 
 */
function onInstall(e) { onOpen(e); }

/**
 * When Doc is opened, create Addon menu entries.
 */
function onOpen(e) { 
  var menu  = DocumentApp.getUi().createAddonMenu()
  .addItem( 'Make sample  version',            'makeSample'  )
  .addItem( 'Make student version',            'makeStudent' )
  .addItem( 'Make teacher version',            'makeTeacher' )
  .addItem( 'Make all versions (may be slow)', 'makeAll'     );

  // if authorized and property set, show advanced / debug features
  if (e && e.authMode != ScriptApp.AuthMode.NONE) { 
    var props = getMergedProperties();
    if ( props[ 'ShowCheck' ] ) {
      menu.addItem( 'Check document'            , 'checkDocument'      )
    }
      menu.addItem( 'Settings...'               , 'showDialogSettings' );
    if ( props[ 'ShowTest' ] ) {
      menu.addSubMenu( DocumentApp.getUi().createMenu("Test")
          .addItem( 'Run unit tests...'         , 'runAll'                   )
          .addItem( 'Show doc attributes...'    , 'showSidebarAttributes'    )
          .addItem( 'Show all properties...'    , 'showSidebarAllProperties' )
          .addItem( 'Remove document header'    , 'removeHeader'             )
          .addItem( 'Remove document footer'    , 'removeFooter'             )
          .addItem( 'Delete doc properties...'  , 'deleteDocProperties'      )
          //.addItem( 'Delete user properties...' , 'deleteUserProperties'     )
          //.addItem( 'Show revision info'        , 'showRevisionInfo'         )
          );
    }
  }

  menu.addToUi();
}

//********************************************************************************
// top-level functions

/**
 * Make copy of document.
 * @param {Document}   oldDoc   Document to copy.    If not specified, default to active document.
 * @param {properties} props    Properties for copy. If not specified, default to merged properties.
 * @param {String}     version  For "Student" (without answers & instructor notes), "Teacher", etc.
 */
function make(oldDoc, props, version) {
  oldDoc      = oldDoc || DocumentApp.getActiveDocument();
  props       = props  || getMergedProperties();
  // get old file, create new file
  var oldFile = DriveApp.getFileById( oldDoc.getId() );
  var newFile = null;
  var newName = getNewName(oldFile.getName(), props, version);
  // if parent folder found, use it and remove old files
  if (         oldFile.getParents().hasNext() ) {
    var fold = oldFile.getParents().next   ();
    if ( props[ 'RemoveOld' ] ) {
      removeFiles(fold, newName);
      removeFiles(fold, newName + ".pdf");
    }
    newFile = oldFile.makeCopy( newName, fold );
  // else (no parent found)
  } else {
    newFile = oldFile.makeCopy( newName );
  }
  var newDoc  = DocumentApp.openById( newFile.getId() );
  var newBody = newDoc.getBody();

  // set header & footer _before_ removing content (before, after, or headings)
  if ( props[ 'ReplaceText' ] ) {
    setHeader   ( newDoc,  props[ 'FormatHeader' ] );
    setFooter   ( newDoc,  props[ 'FormatFooter' ] );
  }
  if        ("Sample" == version) {
    removeAround( newBody, props[ 'SampleStart'  ], props[ 'SampleStop'  ] );
    clearHeading( newBody, props[ 'ClearHeading' ] );
  } else if ("Student" == version) {
    removeAround( newBody, props[ 'StudentStart' ], props[ 'StudentStop' ] );
    clearHeading( newBody, props[ 'ClearHeading' ] );
  } else if ("Teacher" == version) {
    removeAround( newBody, props[ 'TeacherStart' ], props[ 'TeacherStop' ] );
  }
  if ( props[ 'ReplaceText' ] ) {
    replaceAllTags( newDoc, newBody );
  }
  // make readonly, save, save as PDF
  if ( props[ 'CreateRO' ] ) { 
    makeReadonly(newDoc); 
  }
  newDoc.saveAndClose();
  if ( newFile.getParents().hasNext() && props[ 'CreatePDF' ] && 
       ( "Sample" == version || "Student" == version || "Teacher" == version ) ) {
       newFile.getParents().next   ().createFile( newFile.getAs("application/pdf") ).setName( newName + ".pdf" );
  }
  newDoc = DocumentApp.openById( newFile.getId() );
  return newDoc;
}

/**
 * Make sample version of document.
 */
function makeSample()  { return make( null, null, "Sample"  ); }
/**
 * Make student version of document.
 */
function makeStudent() { return make( null, null, "Student" ); }
/**
 * Make teacher version of document.
 */
function makeTeacher() { return make( null, null, "Teacher" ); }
/**
 * Make all versions of document.
 */
function makeAll() { return [ makeSample(), makeStudent(), makeTeacher() ]; }

// EXAMPLE of make() with custom settings 
//function makeDemo() {
//  var props = { 
//    'StudentStart'   :   "{{END STU}}",
//    'StudentStop'    : "{{START STU}}",
//    'StudentSuffix'  : " (Student)",    
//     'FormatHeader'  : "",
//     'FormatFooter'  : "\u00A9 Clif Kussmaul (http://cspogil.org)",
//  };
//  return make( null, props, "Student" );
//}

// FUTURE: get revisions from link & extract useful data
// FUTURE: figure out authorization

/**
 * Show revision info for document.
 * @param {Document}   doc Document to analyze. If not specified, default to active document.
 */
function showRevisionInfo(doc) {
  doc      = doc || DocumentApp.getActiveDocument();
  var id   = doc.getId();
  var revs = Drive.Revisions.list( id );
  var html = "<h1>History for " + doc.getName() + "</h1>\n";
  if (revs.items && revs.items.length > 0) {
    html += "<p># revisions= " + revs.items.length + "</p>\n";
    html += "<ol>\n";
    for (var i=0; i<revs.items.length; i++) {
      var rev = revs.items[i];
// NOTE: fileSize seems to be undefined (maybe for Docs & Sheets, not other file types)
//      var resp = UrlFetchApp.fetch(rev.selfLink);
      html += "<li>Date: " + new Date( rev.modifiedDate ).toLocaleString() + 
                 " By:   " +           rev.lastModifyingUserName + 
                 " Size: " +           rev.fileSize + 
                 " At:   " +           rev.selfLink + 
//                 " Len:  " +           resp.getContentText().length +
          "</li>\n";
    }
    html += "</ol>\n";
  } else {
    html += "<p># revisions= none</p>\n";
  }
  showSidebar("Revisions for " + doc.getName(), html);
}

//********************************************************************************
// check functions

// FUTURE: check colors/fonts, margins & formatting
// FUTURE: settings to choose/config relevant issues to check, or select issue from submenu 
// FUTURE: (not yet in Google API): page numbers & # of pages

// TODO: rename to odoc/doc/body or doc/ndoc/nbod

/**
 * Check document for potential problems.
 * @param {Document}   doc    Document to check.    If not specified, default to active document.
 * @param {properties} props  Properties for check. If not specified, default to merged properties.
 */
function checkDocument(doc, props) {
  doc      = doc   || DocumentApp.getActiveDocument();
  props    = props || getMergedProperties();
  var ndoc = make(doc, props, "Check");
  var body = ndoc.getBody();
  var str  = body.getText();
  // remove tags for sample/student/teacher versions
  var tagList = [ 
    [ props[  'SampleStart' ].slice(2,-2), "", null ], [ props[  'SampleStop' ].slice(2,-2), "", null ],
    [ props[ 'StudentStart' ].slice(2,-2), "", null ], [ props[ 'StudentStop' ].slice(2,-2), "", null ],
    [ props[ 'TeacherStart' ].slice(2,-2), "", null ], [ props[ 'TeacherStop' ].slice(2,-2), "", null ],
  ];
  replaceTextTags(ndoc, body, tagList);
  var html = "<ul>\n"
           + checkSection_   ( ndoc.getHeader(), "Header" )
           + checkSection_   ( ndoc.getFooter(), "Footer" )
           + checkLabels     ( str,  props )
           + checkHeadings   ( body, props[  'ClearHeading'] )
           + checkImages     ( body )
           + checkText       ( body, str, props )
           + checkUnknownTags( body )
           + "</ul>\n";
  showSidebar("Checkup Results", html, "Repeat", "checkDocument();" );
}

/**
 * Check document body for all headings, report counts and how many will be removed. (See CodeTest.gs)
 * @param  {Body}   body          Body to check for given heading.
 * @param  {string} clearHeading  Heading to be cleared.
 * @return {string}               HTML list item with details.
 */
function checkHeadings(body, clearedHeading) {
  if (null == body || null == clearedHeading) return "";
  var html = "<li>Headings</li><ul>\n";
  for ( var heading in HEADINGS ) { // for..in.. only works for properties, NOT for arrays
    var count = countHeading( body, heading );
    if ( count > 0 || heading == clearedHeading ) {
      html += "<li><i>" + heading + "</i>: " + count
           + ( heading == clearedHeading ? " (will be removed)" : " " ) + "</li>\n";
    }
  }
  html += "</ul>\n";
  return html;
}

/**
 * Check images for potential problems.
 * @param  {Body}   body  Body to check for images.
 * @return {string}       HTML list item(s) with details.
 */
function checkImages(body) {
  if (null == body) return "";
  var html    = "";
  var imgList = body.getImages();
  for ( var i=0; i<imgList.length; i++ ) {
    var blob = imgList[i].getBlob(); 
    var len  = Math.round(blob.getBytes().length / 1024);
    var name = blob.getName() || "???";
    if ( 10 < len ) {
      html  += "<li>Image (" + name + ") is " + len + " KB. Reduce size.</li>\n";
    }
  }
  return html;
}

/** 
 * Check for labels in text string.
 * @param  {string}     str    Text string to check.
 * @param  {properties} props  Properties for check. If not specified, default to merged properties.
 * @return {string}            HTML list item with details.
 */
function checkLabels(str, props) {
  if (null == str || null == props) return "";
  var html = "<li>Labels</li><ul>\n"
           + checkLabelStart ( str,  props[ 'TeacherStart'  ] )
           + checkLabelStart ( str,  props[ 'StudentStart'  ] )
           + checkLabelStart ( str,  props[  'SampleStart'  ] )
           + checkLabelStop  ( str,  props[  'SampleStop'   ] )
           + checkLabelStop  ( str,  props[ 'StudentStop'   ] )
           + checkLabelStop  ( str,  props[ 'TeacherStop'   ] )
           + "</ul>\n";
  return html;
}

/** 
 * Check for "start label". Find last  index, remove from start of text  to end of label. (See CodeTest.gs)
 * @param  {string} str    Text string to check.
 * @param  {string} label  Label string to find.
 * @return {string}        HTML list item with details.
 */
function checkLabelStart(str, label) {
  if (null == str || null == label) return "";
  return checkLabel_( str, label, 0, str.lastIndexOf(label) + label.length );
}

/**
 * Check for "stop label". Find first index, remove from start of label to end of text. (See CodeTest.gs)
 * @param  {string} str    Text string to check.
 * @param  {string} label  Label string to find.
 * @return {string}        HTML list item with details.
 */
function checkLabelStop(str, label) {
  if (null == str || null == label) return "";
  return checkLabel_( str, label,    str.indexOf    (label),   str.length );
}

/**
 * Check for label in text string (helper function for checkLabelAfter and checkLabelBefore).
 * @param  {string} str    Text string to check.
 * @param  {string} label  Label string to find.
 * @return {string}        HTML list item with details.
 * @private
 */
function checkLabel_(str, label, fr, to) {
  if (null == str || null == label) return "";
  var count = countText(str, label);
  var html  = "<li><i>" + label + "</i>: " + count + " ";
  switch (count) {
    case 0:  html += "No text will be removed.</li>\n"; return html;
    case 1:                                             break;
    default: html += "<b>Remove duplicates.</b> "     ; break;
  }
  html     += "Chars " + fr + " to " + to + " (" + Math.round(100*(to-fr)/str.length) + "%) will be removed.</li>\n";
  return html;
}

/**
 * Check readability - returns nothing if string is too short. (See CodeTest.gs)
 * @param  {string}  str     Text string to check.
 * @param  {number=} clLim   Colman-Liau    limit (report higher values).
 * @param  {number=} fkLim   Flesch-Kincaid limit (report higher values).
 * @param  {number=} feLim   Flesch Ease    limit (report lower  values).
 * @return {string}          HTML list item(s) with details.
 */
function checkReadability(str, clLim, fkLim, feLim ) {
  // return nothing for short text strings, since readability measures are less useful
  if (str.length < 60) return "";
  // precompute counts rather then recompute for each measure
  var s  = getSentences    ( str ).length;
  var w  = getWords        ( str ).length;
  var l  = getAlphanumerics( str ).length;
  var cl = getReadability  ("ColemanLiau"  , str, s,w,l );
  var fk = getReadability  ("FleschKincaid", str, s,w,l );
  var fe = getReadability  ("FleschEase"   , str, s,w,l );
  var html = "";
  if      (   ! clLim) { html += "<li>Coleman-Liau   = " + cl                  + "  </li>\n"; }
  else if (cl > clLim) { html += "<li>Coleman-Liau   = " + cl  + " (>" + clLim + ") </li>\n"; }
  if      (   ! fkLim) { html += "<li>Flesch-Kincaid = " + fk                  + "  </li>\n"; }
  else if (fk > fkLim) { html += "<li>Flesch-Kincaid = " + fk  + " (>" + fkLim + ") </li>\n"; }
  if      (   ! feLim) { html += "<li>Flesch Ease    = " + fe                  + "  </li>\n"; }
  else if (fe < feLim) { html += "<li>Flesch Ease    = " + fe  + " (<" + feLim + ") </li>\n"; }
  return html;
}

/**
 * Check section (header or footer).
 * @param  {Section} section  Section to check.
 * @param  {string}  label    Label for section ("header" or "footer").
 * @return {string}           HTML list item with details.
 * @private
 */
function checkSection_(section, label) {
  if (null == label) return "";
  var html = "";
  if (null == section) {
    html  += "<li>" + label + " doesn't exist.</li>\n";
  } else {
    var lines = countLines( section.getText() );
    if (lines > 1) { html += "<li>" + label + " has " + lines + " lines.</li>\n"; }
  }
  return html;
}

// FUTURE: check readability, etc. for longer blocks of texts - combine <p> into sections by <h*>?

/**
 * Check text.
 * @param  {Body}       body   Body to check for given heading.
 * @param  {string}     str    Text string to check.
 * @param  {properties} props  Properties for check.
 * @return {string}            HTML list of item(s) with details.
 */
function checkText(body, str, props) {
  if (null == body || null == str || null == props) return "";
  var html     = "<li>Text</li><ul>\n" + checkReadability(str);
  var result   = null;
  while ( result = body.findElement(DocumentApp.ElementType.TEXT, result ) ) {
    var textn  = result.getElement().asText();
    var strn   = textn.getText();
    var htmln  = checkReadability(strn, props['ReadabilityCL'], props['ReadabilityFK'], props['ReadabilityFE'] );
    var size   = textn.getFontSize();
    if ( 5 < strn.length && null != size && size < 10) { htmln  += "<li>Font size = " + size + " (<10) </li>\n"; }
    // generate HTML iff there are interesting results
    if ("" != htmln) { 
      textn.setBackgroundColor("#FFCCCC");
      html += "<li>In <u>" + strn.substring( 0, Math.min( 20, strn.length ) ) + "</u> (" + strn.length + " chars):</li><ul>\n" + htmln + "</ul>\n";
    }
  }
  html      += "</ul>\n";
  return html;
}

/**
 * In document section, check for unknown tags. 
 * @param  {Section}  section  Section  to change. If null, do nothing.
 * @return {string}            HTML list of item(s) with details.
 */
function checkUnknownTags(section) {
  if ( null == section ) return 0;
  var html   = "<li>Unknown tags: <br/>";
  var result = null;
  while (result = section.findText( "{{.*?}}", result )) {
    var elem = result.getElement();
    var tag  =  elem.asText().getText().substring( result.getStartOffset()        + 2,
                                                   result.getEndOffsetInclusive() - 1).trim();
    html += "{{" + tag + "}} ";
  }
  html += "</li>";
  return html;
}


//********************************************************************************
// document functions

/**
 * Clear given heading from body (e.g. to remove answers & facilitator notes). (See CodeTest.gs)
 * @param  {Body}     body     Body to check for given heading.
 * @param  {string}   heading  Heading to search for.
 * @param  {boolean=} clear    true -> clear heading; false -> do not clear heading
 * @result {number}            Number of headings removed.
 */
function clearHeading(body, heading, clear) {
  if (null == body) return 0; 
  return clearHeadingList_( body.getParagraphs(), heading, clear)
       + clearHeadingList_( body.getListItems() , heading, clear);  
}

/**
 * Clear given heading from given list  (e.g. to remove answers & facilitator notes). (See CodeTest.gs)
 * @param  {Array}    list     List of objects to check for given heading.
 * @param  {string}   heading  Heading to search for.
 * @param  {boolean=} clear    true -> clear heading; false -> do not clear heading
 * @result {number}            Number of headings removed.
 */
function clearHeadingList_(list, heading, clear) {
  if (null == clear) clear = true; // clear can be null, false, or true
  if (null == list) return 0;
  if (null == ( heading = HEADINGS[heading]) ) {
    Logger.log("HEADINGS[" + heading + "] is null");
    return 0;
  }
  var count = 0;
  for ( var i=0; i<list.length; i++ ) {
    var item = list[i];
    if (heading == item.getHeading()) {
      count++;
      if (clear) {
        var lines = countLines( item.getText() );
        item.setHeading( DocumentApp.ParagraphHeading.NORMAL );
        if (lines > 2) {
          item.setText( repeat( "\r", lines - 1 ) );
        } else {
          item.clear();
        }
      }
    }
  }
  return count;
}  

// TODO: more efficient: define countHeadings() to loop once through paragraphs, return list of headings & counts

/**
 * Count given heading in body. (See CodeTest.gs)
 * @param  {Body}     body     Body to check for given heading.
 * @param  {string}   heading  Heading to search for.
 * @result {number}            Number of headings removed.
 */
function countHeading(body, heading) { return clearHeading(body, heading, false); }

/** 
 * Count lines in text. (See CodeTest.gs)
 * (Assume 80 chars/line unless there are many explicit line breaks.)
 * @param  {string}     str    Text string to check.
 * @result {number}            Number of lines.
 */
function countLines(str) {
  var charsPerLine = 80;
  return (null == str) ? 0 : Math.max( Math.round( str.length/charsPerLine ), str.split(/\r\n|\r|\n/).length ); 
}

/**
 * Count page breaks in section. (See CodeTest.gs)
 * @param  {Section} section  Section to check.
 * @result {number}           Number of page breaks.
 */
function countPageBreak(section) {
  if (null == section || null == section.findElement ) return 0;
  var count   = 0;
  var result  = null;
  while ( result = section.findElement( DocumentApp.ElementType.PAGE_BREAK, result ) ) { count++; }
  return count;
}

/**
 * Count number of times text string contains label. (See CodeTest.gs)
 * @param  {string} str    Text string to check.
 * @param  {string} label  Label string to find.
 * @result {number}        Number of occurrences.
 */
function countText(str,label) {
  if (null == str || null == label) return 0;
  var count =  0;
  var index = -1;
  while ( -1 != ( index = str.indexOf(label, ++index) ) ) { count++; }
  return count;
}

function getNewName(oldName, props, version) {
  var newName = oldName.replace(            props[ 'AuthorSuffix'], "" )
                + (("Sample"  == version) ? props[ 'SampleSuffix'] : 
                   ("Student" == version) ? props['StudentSuffix'] : 
                   ("Teacher" == version) ? props['TeacherSuffix'] : (" - " + version) );
  newName = newName.replace(/\s+/g, " ");
  return newName;
}

// NOTE: 4-5 sentences or 200-500 words (?) needed for significant results - see limit checked in checkReadability()
// NOTE: Gunning-Fog Index uses "complex words" and ignores proper nouns, jargon, compound words, or suffixes.

/**
 * Get readability measure of text string.
 * @param  {string}  measure  Readability measure ("ColemanLiau", "FleshEase", or "FleschKincaid")
 * @param  {string}  str      Text string to check.
 * @param  {number=} sents    Number of sentences in text string.
 * @param  {number=} words    Number of words     in text string.
 * @param  {number=} letts    Number of letters   in text string.
 * @return {number}           Readability value.
 */
function getReadability(measure, str, sents, words, letts) {
  if (null == measure) return 0; 
  if (null == str    ) str = "";
  sents = sents || getSentences    (str).length;
  words = words || getWords        (str).length;
  letts = letts || getAlphanumerics(str).length;
  var syls = letts / 3;
  switch (measure) {
    case 'ColemanLiau'  : return Math.round( 100 *   ( 5.88 * letts / words - 29.6 * sents / words - 15.8 ) ) / 100;
    case 'FleschEase'   : return Math.round(206.835 - 1.015 * words / sents - 84.6 * syls / words);
    case 'FleschKincaid': return Math.round( 100 *   ( 0.39 * words / sents + 11.8 * syls / words - 15.59 ) ) / 100;
    default             : return 0;
  }
}

/**
 * Get alphanumeric characters in text string. (See CodeTest.gs)
 * @param  {string}         str  Text string to check.
 * @return {<string>}            Alphanumeric characters.
 */
function getAlphanumerics(str) {
  return (null == str) ?  ""  : str.replace(/\W+/g,""); // remove non-alphanumeric characters
}

/**
 * Get array of sentences in text string. (See CodeTest.gs)
 * @param  {string}         str  Text string to check.
 * @return {Array.<string>}      Array of sentences.
 */
function getSentences(str) {
  return (null == str) ? [""] : str.split(/\W*(?:\.|\?|\!)\W*/); // split by .?!
}

/**
 * Get array of words in text string. (See CodeTest.gs)
 * @param  {string}         str  Text string to check.
 * @return {Array.<string>}      Array of words.
 */
function getWords(str) {
  return (null == str) ? [""] : str.split(/\W+/); // split by 1+ non-alphanumeric characters
}

/**
 * Get title of section. (See CodeTest.gs)
 * @param  {Section} section  Section to check.
 * @return {string}           Title of section.
 */
function getTitle(section) {
  if (null == section) return "";
  var parList = section.getParagraphs();
  for ( var i=0; i<parList.length; i++ ) {
    var par = parList[i];
    if ( DocumentApp.ParagraphHeading.TITLE == par.getHeading() ) { return par.getText(); }
  }
  return "";
}

/**
 * True if element contains string. (See CodeTest.gs)
 * @param  {Element}  element  Element to search in.
 * @param  {string}   str      Text string to search for.
 * @return {boolean}           true if found, false otherwise.
 */
function hasText(element, str) {
  return ( null != element          ) &&
         ( null != element.getText  ) &&
         ( -1   != element.getText().indexOf( str ) );
}

/**
 * Make doc readonly - change all editors to viewers.
 * @param {Document}   doc    Document to check.    If not specified, do nothing.
 */
function makeReadonly(doc) {
  if (null == doc) return;
  var eds = doc.getEditors();
  for (var i in eds) { doc.removeEditor( eds[i] ).addViewer( eds[i] ); }
}

/**
 * Remove body content around paragraphs with start and end tags. (See CodeTest.gs)
 * @param  {Body}    body      Document body to be modified.
 * @param  {string}  startTag  start tag  to look for in body
 * @param  {string}   stopTag  stop  tag  to look for in body
 * @return {number}            Number of paragraphs removed.
 */
function removeAround(body, startTag, stopTag) {
  if (null == body || null == body.getNumChildren) return 0;
  if (null == startTag || "" == startTag || null == stopTag || "" == stopTag) return 0;
  var end        = body.getNumChildren()-1;
  var remFirst  = false;
  var result    = 0;
  // scan backwards - if stop, remove stop to end; if start or stop, reset end
  for (var i=end; i>=0; i--) {
    var child = body.getChild(i);
    if        ( hasText( child, startTag ) && ! remFirst ) {
      end       = i;
      remFirst  = true;
    } else if ( hasText( child,  stopTag ) ) {
      result += removeChildren_(body, i, end);
      end       =  i-1;
      remFirst  = false;
    }
  }
  // if last, delete from 0 to end
  if (remFirst) { result += removeChildren_(body, 0, end); }
  return result;
}

/**
 * Helper method for removeAround.
 * @param  {Body}   body  Document body to be modified.
 * @param  {number} min   index of first child to remove
 * @param  {number} min   index of last  child to remove
 * @return {number}       Number of paragraphs removed.
 * @private
 */
function removeChildren_(body, mini, maxi) {
  var count = 0;
  if ( body.getNumChildren() - 1 == maxi ) body.appendParagraph(""); 
  for (var j=maxi; j >= mini; j--) {
    body.getChild(j).removeFromParent();
    count++;
  }
  return count;
}

// TODO: how to detect problems here? does setTrashed throw exceptions? (not documented)
//       maybe something like: try {setTrashed(); } catch (e) { showAlert(e); }

/**
 * Remove files with given name from given folder (filenames are not unique in Google Drive).
 * @param  {Folder} folder    Folder from which to remove files.
 * @param  {string} filename  Name of file(s) to be removed.
 * @return {number}           Number of files removed.
 */
function removeFiles(folder, filename) {
  if (null == folder || null == folder.getFilesByName || null == filename) return 0;
  var count = 0;
  var list  = folder.getFilesByName(filename);
  while ( list.hasNext() ) { list.next().setTrashed(true); count++; }
  return count;
}

/**
 * Remove document footer (called from menu).
 * @param  {Document=} doc      Document to check. If not specified, default to active document.
 * @param  {boolean=}  confirm  If true, skip confirmation alert.
 * @return {boolean}            True if removed, false otherwise
 */
function removeFooter(doc, confirm) {
  doc = doc || DocumentApp.getActiveDocument();
  return removeSection_( doc.getFooter(), confirm );
}

/**
 * Remove document header (called from menu).
 * @param  {Document=} doc      Document to check. If not specified, default to active document.
 * @param  {boolean=}  confirm  If true, skip confirmation alert.
 * @return {boolean}            True if removed, false otherwise
 */
function removeHeader(doc, confirm) {
  doc = doc || DocumentApp.getActiveDocument();
  return removeSection_( doc.getHeader(), confirm );
}

/**
 * Helper function for removeFooter() and removeHeader().
 * @param  {Section}  section  Section to remove. If not specified, do nothing.
 * @param  {boolean=} confirm  If true, skip confirmation alert.
 * @return {boolean}           True if removed, false otherwise
 * @private
 */
function removeSection_(section, confirm) {
  if (null == section) return false;
  confirm = confirm || false;
  if (confirm || DocumentApp.getUi().Button.OK == showAlert("Confirm removal","Really remove?")) {
    section.removeFromParent();
    return true;  
  }
  return false;
}

/**
 * Repeat given string given number of times. (See CodeTest.gs)
 * @param  {string} str  String to repeat.
 * @param  {number} num  Number of times to repeat.
 * @return {string}
 */
function repeat(str, num) { return new Array( num+1 ).join(str); }


// NOTE: Body has all Section methods + page break & page/margin size
// NOTE: Paragraph has InlineImage not Image

/**
 * In given section, replace given tag with image and optional link. (See CodeTest.gs)
 * @param  {Section} section  Section to change.       If null         , do nothing.
 * @param  {string}  tag      Tag string to find.      If null or empty, do nothing.
 * @param  {string}  image    URL for image to insert. If null         , do nothing.
 * @param  {string=} link     URL for link  to insert.
 * @return {number}           Number of tags replaced.
 */
function replaceImageTag(section, tag, image, link) {
  if ( null == section || null == section.findElement || 
       null == tag     || ""   == tag || 
       null == image ) return 0;  
  // if no tag, do nothing
  if ( null == section.findText(tag) ) return 0;
  var blob    = UrlFetchApp.fetch( image ).getBlob();
  var count   = 0;
  var parList = section.getParagraphs();
  // loop through paragraphs in section
  for ( var i=0; i<parList.length; i++ ) {
    var par     = parList[i];
    var resultP = null;
    // find text elements in paragraphs
    while (resultP = par.findElement( DocumentApp.ElementType.TEXT, resultP )) {
      var text    = resultP.getElement().asText();
      var resultT = null;
      // find tag in text element - repeated tags will be pushed into new sibling text elements
      if (resultT = text.findText( tag, resultT )) {
        var k1   = resultT.getStartOffset();
        var k2   = resultT.getEndOffsetInclusive();
        var k3   = text.getText().length - 1;
        var next = text.copy();
        text.deleteText( k1, k3 ); // text before tag - remove start of tag to end
        next.deleteText(  0, k2 ); // text after  tag - remove start to end of tag
        var rent = text.getParent();
        var j    = rent.getChildIndex    (      text );
        var img  = rent.insertInlineImage( j+1, blob );
                   rent.insertText       ( j+2, next );
        if (null != link) { img.setLinkUrl( link ); }
        count++;
      }
    }
  }
  return count;
}

/**
 * In given section, replace given tag with text and optional link. (See CodeTest.gs)
 * @param  {Section} section  Section to change.  If null         , do nothing.
 * @param  {string}  tag      Tag string to find. If null or empty, do nothing.
 * @param  {string}  repl     Text to insert.     If null         , do nothing.
 * @param  {string=} link     URL for link to insert.
 * @return {number}           Number of elements with tags replaced (<= actual count).
 */
function replaceTextTag(section, tag, repl, link) {
  if ( null == section || null == section.findText || 
       null == tag     || ""   == tag || 
       null == repl ) return 0;
  var count  = 0;
  var result = null;
  // always search from start, since element changes inside loop
  // replaceText() replaces all occurrences
  while (result = section.findText( tag, result )) {
    var pos  = result.getStartOffset();
    var text = result.getElement().asText().replaceText( tag, repl );
    if (null != link) { text.setLinkUrl( pos, pos + repl.length, link ); }
    count++;
  }
  return count;
}

// FUTURE: user-defined tags, in Settings 
// FUTURE: ? more tags - TAB
// FUTURE: ? more tags - not (yet) supported by Google Apps Script API: USER, EMAIL, PAGE

/**
 * In section of document, replace all known tags. 
 * @param  {Document} doc      Document to change. If null, do nothing.
 * @param  {Section}  section  Section  to change. If null, do nothing.
 * @return {number}            Number of tags replaced.
 */
function replaceAllTags(doc, section) {
  if (null == doc || null == section) return 0;
  // order matters - replaceText() tries to clean up replaceImportTags
  var impCount = replaceImportTags   ( doc, section );
  var txtCount = replaceTextTags     ( doc, section );  
  var imgCount = replaceImageTags    ( doc, section );
  var attList  = replaceAttributeTags( doc, section );
  section.setAttributes( attList );
  return txtCount + imgCount + Object.keys(attList).length;
}

// TODO: are these tags ever needed or used? can they be removed?

/**
 * In document section, replace all known attribute tags. 
 * @param  {Document} doc        Document to change. If null, do nothing.
 * @param  {Section}  section    Section  to change. If null, do nothing.
 * @param  {Array=}   tagList    Array of tags to replace.
 * @return {Object.<string,*>}   Set of attributes to change.
 */
function replaceAttributeTags(doc, section, tagList) {
  var attrs = {};
  if (null == doc || null == section) return attrs;
  tagList = tagList || [
    [ 'BOLD'    , DocumentApp.Attribute.BOLD                , true ],
    [ 'ITALIC'  , DocumentApp.Attribute.ITALIC              , true ],
    [ 'UNDER'   , DocumentApp.Attribute.UNDERLINE           , true ],
    [ 'STRIKE'  , DocumentApp.Attribute.STRIKETHROUGH       , true ],
    [ 'FONT10'  , DocumentApp.Attribute.FONT_SIZE           , 10 ],
    [ 'FONT12'  , DocumentApp.Attribute.FONT_SIZE           , 12 ],
    [ 'FONT14'  , DocumentApp.Attribute.FONT_SIZE           , 14 ],
    [ 'ARIAL'   , DocumentApp.Attribute.FONT_FAMILY         , DocumentApp.FontFamily.ARIAL            ],
    [ 'COURIER' , DocumentApp.Attribute.FONT_FAMILY         , DocumentApp.FontFamily.COURIER_NEW      ],
    [ 'GEORGIA' , DocumentApp.Attribute.FONT_FAMILY         , DocumentApp.FontFamily.GEORGIA          ],
    [ 'TIMES'   , DocumentApp.Attribute.FONT_FAMILY         , DocumentApp.FontFamily.TIMES_NEW_ROMAN  ],
    // TODO: fix alignment - may need to be applied at a different level
    [ 'CENTER'  , DocumentApp.Attribute.HORIZONTAL_ALIGNMENT, DocumentApp.HorizontalAlignment.CENTER  ],
    [ 'JUSTIFY' , DocumentApp.Attribute.HORIZONTAL_ALIGNMENT, DocumentApp.HorizontalAlignment.JUSTIFY ],
    [ 'LEFT'    , DocumentApp.Attribute.HORIZONTAL_ALIGNMENT, DocumentApp.HorizontalAlignment.LEFT    ],
    [ 'RIGHT'   , DocumentApp.Attribute.HORIZONTAL_ALIGNMENT, DocumentApp.HorizontalAlignment.RIGHT   ],
  ];
  for (var i in tagList) {
    if ( 0 < replaceTextTag( section, '{{' + tagList[i][0] + '}}', "" ) ) { attrs[ tagList[i][1] ] = tagList[i][2]; }
  }
  return attrs;
}

/**
 * In document section, replace all known image tags. 
 * @param  {Document} doc      Document to change. If null, do nothing.
 * @param  {Section}  section  Section  to change. If null, do nothing.
 * @param  {Array=}   tagList  Array of tags to replace.
 * @return {number}            Number of tags replaced.
 */
function replaceImageTags(doc, section, tagList) {
  if (null == doc || null == section) return null;
  tagList = tagList || [
    [ 'CC-BY-IMG'      , "https://i.creativecommons.org/l/by/4.0/80x15.png"
                       , "http://creativecommons.org/licenses/by/4.0/"            ],
    [ 'CC-BY-NC-IMG'   , "https://i.creativecommons.org/l/by-nc/4.0/80x15.png"
                       , "http://creativecommons.org/licenses/by-nc/4.0/"         ],
    [ 'CC-BY-NC-SA-IMG', "https://i.creativecommons.org/l/by-nc-sa/4.0/80x15.png"
                       , "http://creativecommons.org/licenses/by-nc-sa/4.0/"      ],
  ];
  var count = 0;
  for (var i in tagList) { 
    count += replaceImageTag( section, '{{' + tagList[i][0] + '}}', tagList[i][1], tagList[i][2] ); 
  }
  return count;
}

// URGENT: ensure that replaceImportTags works with both ID and URL - swapping openBy calls doesn't work
// TODO: replaceImportTags() should handle multiple tags {{A}}{{B}}, but doesn't - see new replaceImageTag()
// - probably need to collect list of results, then process back to front

/**
 * In document section, replace IMPORT tags. 
 * @param  {Document} doc      Document to change. If null, do nothing.
 * @param  {Section}  section  Section  to change. If null, do nothing.
 * @return {number}            Number of tags replaced.
 */
function replaceImportTags(doc, section) {
  if ( null == doc || null == section ) return 0;
  var count  = 0;
  var result = null;
  while (result = section.findText( "{{IMPORT.*?}}", result )) {
    var elem = result.getElement();
    // elem is Text, parent is Paragraph, parent is Body
    var ind  =  elem.getParent().getParent().getChildIndex( elem.getParent() );
    var file =  elem.asText().getText().substring( result.getStartOffset()        + 9,
                                                   result.getEndOffsetInclusive() - 1).trim();
    try {
      var idoc   = DocumentApp.openById(file) || DocumentApp.openByUrl(file);
    } catch (err) {
      showAlert("Import error", "Couldn't import " + file + ". Check logs.");
      Logger.log("Couldn't open " + file + ": " + err.message); continue;
    }
    if (null == idoc) continue;
    var header = idoc.getHeader(); if (null != header) { setHeader(doc, header.getText() ); }
    var footer = idoc.getFooter(); if (null != footer) { setFooter(doc, footer.getText() ); }
    var ibod   = idoc.getBody();
    for (var i=0; i<ibod.getNumChildren(); i++) {
      var e = ibod.getChild(i).copy();
      switch ( e.getType() ) {
        case DocumentApp.ElementType.HORIZONTAL_RULE: section.insertHorizontalRule(ind+i   ); break;
        case DocumentApp.ElementType.INLINE_IMAGE   : section.insertImage         (ind+i, e); break;
          // TODO: list items are not formatted correctly
        case DocumentApp.ElementType.LIST_ITEM      : section.insertListItem      (ind+i, e)
                                                                .setGlyphType   ( e.getGlyphType   () )
                                                                .setListId      ( e )
                                                                .setNestingLevel( e.getNestingLevel() ); break;
        case DocumentApp.ElementType.PAGE_BREAK     : section.insertPageBreak     (ind+i   ); break;
        case DocumentApp.ElementType.PARAGRAPH      : section.insertParagraph     (ind+i, e); break;
        case DocumentApp.ElementType.TABLE          : section.insertTable         (ind+i, e); break;
        default: showAlert("Import", "ElementType=" + e.getType() ); break;
      }
    }
    elem.getParent().removeFromParent();
    count++;
  }
  return count;    
}

/**
 * In document section, replace all known text tags. 
 * @param  {Document} doc      Document to change. If null, do nothing.
 * @param  {Section}  section  Section  to change. If null, do nothing.
 * @param  {Array=}   tagList  Array of tags to replace.
 * @return {number}            Number of tags replaced.
 */
function replaceTextTags(doc, section, tagList) {
  if (null == doc || null == section) return null;
  tagList = tagList || [
    [ 'DATE'           , Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd")  , null ],
    [ 'YEAR'           , Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy"      )  , null ],
    [ 'NAME'           , doc.getName()                                                                , null ],
    [ 'TITLE'          , getTitle( doc.getBody() )                                                    , null ],
    [ 'NEWLINE'        , "\r"                                                                         , null ],
    [ 'CC-BY'          , "{{CC-BY-IMG}} This work is licensed under a {{CC-BY-TXT}}."                 , null ],
    [ 'CC-BY-TXT'      , "Creative Commons Attribution 4.0 International License"
                       , "http://creativecommons.org/licenses/by/4.0/" ],
    [ 'CC-BY-NC'       , "{{CC-BY-NC-IMG}} This work is licensed under a {{CC-BY-NC-TXT}}."           , null ],
    [ 'CC-BY-NC-TXT'   , "Creative Commons Attribution-NonCommercial 4.0 International License"
                       , "http://creativecommons.org/licenses/by-nc/4.0/" ],
    [ 'CC-BY-NC-SA'    , "{{CC-BY-NC-SA-IMG}} This work is licensed under a {{CC-BY-NC-SA-TXT}}."     , null ],
    [ 'CC-BY-NC-SA-TXT', "Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International License"
                       , "http://creativecommons.org/licenses/by-nc-sa/4.0/" ],
    [ 'CS-POGIL'       , "The CS-POGIL Project", "http://cspogil.org" ],
    [    'POGIL'       ,    "The POGIL Project",    "http:/pogil.org" ],
  ];
  var count = 0;
  for (var i in tagList) { 
    count += replaceTextTag( section, '{{' + tagList[i][0] + '}}', tagList[i][1], tagList[i][2] ); 
  }
  return count;
}


/**
 * Set document footer (if null), then replace tags.
 * @param  {Document} doc   Document to change. If null, do nothing.
 * @param  {string}   str   Text string for footer (may include tags).
 * @return {number}         Number of tags replaced.
 */
function setFooter(doc, str) {
  if (null == doc) return 0;
  var footer = doc.getFooter()  ||  ( null != str && "null" != str && doc.addFooter().setText( str ) )  || null;
  return replaceAllTags( doc, footer );
}

/**
 * Set document header (if null), then replace tags.
 * @param  {Document} doc   Document to change. If null, do nothing.
 * @param  {string}   str   Text string for header (may include tags).
 * @return {number}         Number of tags replaced.
 */
function setHeader(doc, str) {
  if (null == doc) return 0;
  var header = doc.getHeader()  ||  ( null != str && "null" != str && doc.addHeader().setText( str ) )  || null;
  return replaceAllTags( doc, header );
}

//********************************************************************************
// properties functions

/**
 * Delete all document properties.
 */
function  deleteDocProperties() { 
  PropertiesService.getDocumentProperties().deleteAllProperties();
  showSidebarAllProperties();
}

/**
 * Delete all user properties.
 */
function deleteUserProperties() { 
  PropertiesService.getUserProperties    ().deleteAllProperties();
  showSidebarAllProperties();
}

/** 
 * Get merged properties.
 */
function getMergedProperties() {
  var sprops = PropertiesService.getScriptProperties();
  var uprops = PropertiesService.getUserProperties();
  var dprops = PropertiesService.getDocumentProperties();
  var  props = {};
  for (var key in sprops.getProperties()) { props[ key ] = sprops.getProperty( key ); }
  for (var key in uprops.getProperties()) { props[ key ] = uprops.getProperty( key ); }
  for (var key in dprops.getProperties()) { props[ key ] = dprops.getProperty( key ); }
  return props;
}

/**
 * Get HTML table rows with given heading from given properties - either type.
 * @param  {string} heading
 * @param  {Object} props
 * @return {string}          HTML
 */
function getPropertyRows(heading, props) {
  var html  = "<tr><th>Key</th><th>" + heading + " Property Value</th></tr>\n";
  var list  = props.getProperties ? props.getProperties()  : props;
  var keys   = Object.getOwnPropertyNames(list).sort();
  if (0 == keys.length) {
    html   += "<tr><td> (none) </td><td> (none) </td></tr>\n"; 
  }
  for (var i in keys) { // for..in.. only works for properties, NOT for arrays
    var val = props.getProperty   ? props.getProperty(keys[i]) : props[keys[i]];
    html   += "<tr><td>" + keys[i] + "</td><td>" + val + "</td></tr>\n"; 
  }
  return html;
} 

/**
 * Set merged properties.
 * @param  {Object} props
 */
function setMergedProperties(props) {
  var sprops = PropertiesService.getScriptProperties();
  var uprops = PropertiesService.getUserProperties();
  var dprops = PropertiesService.getDocumentProperties();
  for (var key in sprops.getProperties()) { // for..in.. only works for properties, NOT for arrays
    Logger.log( "key: " + key + " val: " + props[key] );

    // if new value matches script value, remove user value & doc value if they exist
    if        ( props[key] == sprops.getProperty(key) ) {
      Logger.log("  in script: " + key);
      if ( null != uprops.getProperty(key) ) { uprops.deleteProperty(key); }
      if ( null != dprops.getProperty(key) ) { dprops.deleteProperty(key); }

    // else if new value matches user value, remove doc value 
    } else if ( props[key] == uprops.getProperty(key) ) {
      Logger.log("  in user: " + key);
      if ( null != dprops.getProperty(key) ) { dprops.deleteProperty(key); }

    // else if new value doesn't match doc value, update doc value
    } else if ( props[key] != dprops.getProperty(key) ) {
      Logger.log("  set in doc: " + key);
      dprops.setProperty(key, props[key]);
    }
  }
}

//********************************************************************************
// UI functions

/**
 * Show alert with title, text, and button.
 * @param  {string=} title
 * @param  {string=} text
 * @param  {string=} button label
 * @return           Result from alert.
 */
function showAlert(title, text, button) {
  var ui = DocumentApp.getUi();
  title  = title  || "(No Alert Title)";
  text   = text   || "(No Alert Text)";
  button = button || ui.ButtonSet.OK_CANCEL;
  return ui.alert( title, text, button );
}

/**
 * Show dialog with title, HTML output, and dimensions.
 * @param  {string=}    title
 * @param  {HTMLOutput} output
 * @param  {number=}    width
 * @param  {number=}    height
 * @private
 */ 
function showDialogOut_(title, out, width, height) {
  if (null == out) return;
  width   = width  || 200;
  height  = height || 200;
  out.setWidth(width).setHeight(height);
  DocumentApp.getUi().showModalDialog(out, title);
}

/**
 * Show dialog with title, HTML, and dimensions.
 * @param  {string=} title
 * @param  {string=} html
 * @param  {number=} width
 * @param  {number=} height
 */
function showDialog(title, html, width, height) {
  title   = title  || "(No Dialog Title)";
  html    = html   || "<p>(No Dialog HTML)</p>";
  var css = '<link rel="stylesheet" href="https://ssl.gstatic.com/docs/script/css/add-ons.css">\n';
  var out = HtmlService.createHtmlOutput(css + html);
  showDialogOut_(title, out, width, height);
}

/**
 * Show dialog with title, content file, and dimensions.
 * @param  {string=} title
 * @param  {string=} file
 * @param  {number=} width
 * @param  {number=} height
 * @return           Result from modelDialog().
 */
function showDialogFromFile(title, file, width, height) {
  width      = width  || 400;
  height     = height || 400;
  var temp   = HtmlService.createTemplateFromFile(file);
  temp.props = getMergedProperties();
  showDialogOut_(title, temp.evaluate(), width, height);
}

/**
 * Show dialog with add-on settings.
 */
function showDialogSettings() { showDialogFromFile('Settings', 'Settings', 600, 600); }

/**
 * Show prompt with title, text, and button.
 * @param  {string=} title
 * @param  {string=} text
 * @param  {string=} button label
 * @return           Result from prompt.
 */
function showPrompt(title, text, button) {
  var ui = DocumentApp.getUi();
  title  = title  || "(No Prompt Title)";
  text   = text   || "(No Prompt Text)";
  button = button || ui.ButtonSet.OK_CANCEL;
  return ui.prompt( title, text, button );
}

// NOTE: Google Scripts ignores sidebar setWidth() as of 2014-02

/**
 * Show sidebar with title, html, and action button.
 * @param  {string=} title
 * @param  {string=} html
 * @param  {string=} action  label for action button
 * @param  {string=} func    Javascript function to run when button is pressed.
 */
function showSidebar(title, html, action, func) {
  title   = title || "(No Title)";
  html    = html  || "<p>(No HTML)</p>";
  if (action && func) {
    html  = "<p><button class='action' id='repeat' onclick='this.disabled=true; google.script.run." + func + "'>" + action + "</button>"
          +    "<button                id='cancel' onclick='google.script.host.close();'                       >Cancel</button></p>"
          + html
  }
  html    = '<link rel="stylesheet" href="https://ssl.gstatic.com/docs/script/css/add-ons.css">\n' 
          + html;
  var out = HtmlService.createHtmlOutput(html).setTitle(title);
  DocumentApp.getUi().showSidebar( out );
}

/**
 * Show sidebar using file and properties.
 * @param  {string} file
 * @param  {Object} props
 */
function showSidebarFromFile(file, props) {
  if (null == file || null == props) return;
  var t = HtmlService.createTemplateFromFile(file);
  t.props = props || getMergedProperties();
  DocumentApp.getUi().showSidebar( t.evaluate().setTitle(file) );
}

/**
 * Show sidebar listing all properties (script, user, doc, merged).
 */
function showSidebarAllProperties() {
  var html = "<table>\n"
           + getPropertyRows( "Script", PropertiesService.getScriptProperties  () )
           + getPropertyRows(   "User", PropertiesService.getUserProperties    () )
           + getPropertyRows(    "Doc", PropertiesService.getDocumentProperties() )
           + getPropertyRows( "Merged",                   getMergedProperties  () )
           + "</table>\n";
  showSidebar( "Properties", html, "Repeat", "showSidebarAllProperties();" );
}

/**
 * Show sidebar listing all document attributes.
 * @param  {Document=} doc  Document to list. If not specified, default to active document.
 */
function showSidebarAttributes(doc) {
  doc        = doc || DocumentApp.getActiveDocument();
  var atts   = doc.getBody().getAttributes();
  var keys   = Object.getOwnPropertyNames(atts).sort();
  var html   = "<table><tr><th>Attribute</th><th>Value</th></td>";
  for (var i in keys) { // for..in.. only works for properties, NOT for arrays
    html += "<tr><td>" + keys[i] + "</td><td>" + atts[keys[i]] + "</td></tr>\n"; 
  } 
  html      += "</table>";
  showSidebar( "Attributes", html, "Repeat", "showSideBarAttributes();" );
}

//********************************************************************************
// end of file
