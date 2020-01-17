/**
 * @preserve CodeTest.gs
 * - part of Classroom Activity Utility (CAU) (Google Doc add-on)
 * - Copyright (c) 2014 Clif Kussmaul, clif@kussmaul.org, clif@cspogil.org
 */

// global variables - so  assert() expressions can access doc & body and append to htmlErr
var doc, body, htmlErr, props;

/**
 * Run all test functions, build lists of failed and passed tests, generate test report.
 */
function runAll() {
  doc           = DocumentApp.create("CAUTest");
  body          = doc.getBody();
  props         = getMergedProperties();
  var htmlFail  = "";
  var htmlPass  = "";
  var names     = Object.getOwnPropertyNames(this).sort(); 
  for (var i in names) {
    var func = this[ names[i] ];
    // skip everything but test functions
    if ( "function" != typeof func || 0 != func.name.indexOf("test") ) continue;
    htmlErr = "";
    try           { body.clear(); func(); 
    } catch (err) { htmlErr += "  <ul><li>ERROR: " + err + "</li></ul>\n";
    }
    var htmlNext  = "<li>" + func.name.substring(4) + "</li>\n";                 
    if ("" == htmlErr) { htmlPass += htmlNext } else { htmlFail += htmlNext + htmlErr; }
  }
  if ("" != htmlFail) { htmlFail = "<h2>Fail</h2><ol>\n" + htmlFail + "</ol>\n"; }
  html  = "<p><button class='action' id='repeat' onclick='this.disabled=true; google.script.run.runAll();'>Rerun Tests</button>\n"
        +    "<button                id='cancel' onclick='google.script.host.close()'                     >Cancel</button>\n"
        + "</p>" + htmlFail + "<h2>Pass</h2>\n<ol>" + htmlPass + "</ol>\n";
  showSidebar("Unit Test Results", html, 300);
  DriveApp.getFileById( doc.getId() ).setTrashed( true );
}

/** 
 * Append variable number of paragraphs to body.
 */
function appendParagraphs_() {
  for (var i=0; i<arguments.length; i++) { body.appendParagraph(arguments[i]); }
  return arguments.length;
}

/**
 * Evaluate expression and report error if actual value != expected value,
 * or if delta is defined and abs(actual - expected) > delta
 */
function assert_(expect, expr, delta) {
  var actual = eval(expr);
  if (! equals_( expect, actual, delta ) ) { 
    htmlErr += "  <ul><li>FAIL: " + expr + " => " + actual + " expected " + expect + "</li></ul>\n";
  }
}

/** 
 * Decide if two values are equal (within error margin).
 * @param  {*}       a  1st value to compare.
 * @param  {*}       b  2nd value to compare.
 * @param  {number=} d  error margin.
 * @return {boolean}    True if equal, false otherwise.
 */
function equals_(a, b, d) {
  if (! (a instanceof Array && b instanceof Array) )  
    return (null == d) ? ( a == b ) : ( Math.abs( a - b ) <= d );
  if ( a.length != b.length )                                    return false;
  for (var i=0; i<a.length; i++) { if (! equals_( a[i], b[i] ) ) return false; }
  return true;
}

// ********************************************************************************
// test data

var fish  = "One fish, two fish, red fish, blue fish.";
var getty = "Four score and seven years ago, our forefathers brought forth on this continent a new nation.";
    
// ********************************************************************************
// individual tests

function testCheckHeadings_() {
  assert_( "<li>Headings</li><ul>\n" + 
             "<li><i>Heading4</i>: 0 (will be removed)</li>\n" + 
             "<li><i>Normal</i>: 1 </li>\n" + 
           "</ul>\n", 'checkHeadings( body, "Heading4" )' );
  body.insertParagraph(0,"XYZ").setHeading( DocumentApp.ParagraphHeading.HEADING4 );
  body.insertParagraph(0,"XYZ").setHeading( DocumentApp.ParagraphHeading.HEADING4 );
  assert_( "<li>Headings</li><ul>\n" + 
             "<li><i>Heading4</i>: 2 (will be removed)</li>\n" + 
             "<li><i>Normal</i>: 1 </li>\n" + 
           "</ul>\n", 'checkHeadings( body, "Heading4" )' );
  assert_( "<li>Headings</li><ul>\n" + 
             "<li><i>Heading3</i>: 0 (will be removed)</li>\n" + 
             "<li><i>Heading4</i>: 2 </li>\n" +
             "<li><i>Normal</i>: 1 </li>\n" + 
           "</ul>\n", 'checkHeadings( body, "Heading3" )' );
}
function testCheckHeadings_Null_() {
  assert_( "", 'checkHeadings( "str", null  )' );
  assert_( "", 'checkHeadings( null , "str" )' );
}

function testCheckImages_Null_() {
  assert_( "", 'checkImages( null )' );
}

function testCheckLabels_() {
  assert_( "<li>Labels</li><ul>\n"
          +   "<li><i>{{STUDENT START}}</i>: 1 Chars 0 to 23 (46%) will be removed.</li>\n"
          +   "<li><i>{{STUDENT STOP}}</i>: 1 Chars 29 to 50 (42%) will be removed.</li>\n"
          + "</ul>\n", 
          'checkLabels( '
          + ' "BEFORE{{STUDENT START}}MIDDLE{{STUDENT STOP}}AFTER", ' 
          + ' { "StudentStart" : "{{STUDENT START}}" , '
          +   ' "StudentStop"  : "{{STUDENT STOP}}"  , } )' );
}
function testCheckLabels_Null_() {
  assert_( "", 'checkLabels( "str", null )' );
  assert_( "", 'checkLabels( null ,  []  )' );
}

function testCheckLabelStart_Found_() {
  assert_( "<li><i>{{LABEL}}</i>: 1 Chars 0 to 15 (75%) will be removed.</li>\n", 'checkLabelStart( "BEFORE{{LABEL}}AFTER", "{{LABEL}}" )' );
}
function testCheckLabelStart_Not_() {
  assert_( "<li><i>{{OTHER}}</i>: 0 No text will be removed.</li>\n", 'checkLabelStart( "BEFORE{{LABEL}}AFTER", "{{OTHER}}" )' );
}
function testCheckLabelStart_Null_() {
  assert_( "", 'checkLabelStart( "str", null  )' );
  assert_( "", 'checkLabelStart( null , "str" )' );
}

function testCheckLabelStop_Found_() {
  assert_( "<li><i>{{LABEL}}</i>: 1 Chars 6 to 20 (70%) will be removed.</li>\n", 'checkLabelStop( "BEFORE{{LABEL}}AFTER", "{{LABEL}}" )' );
}
function testCheckLabelStop_Not_() {
  assert_( "<li><i>{{OTHER}}</i>: 0 No text will be removed.</li>\n", 'checkLabelStop( "BEFORE{{LABEL}}AFTER", "{{OTHER}}" )' );
}
function testCheckLabelStop_Null_() {
  assert_( "", 'checkLabelStop( "str", null  )' );
  assert_( "", 'checkLabelStop( null , "str" )' );
}

function testCheckText_Null_() {
  assert_( "", 'checkText( "str", "str", null  )' );
  assert_( "", 'checkText( "str", null , "str" )' );
  assert_( "", 'checkText( null , "str", "str" )' );
} 

function testCheckReadability_() {
  var ansFish0 = "<li>Coleman-Liau   = -0.07  </li>\n"
               + "<li>Flesch-Kincaid = 0.54  </li>\n"
               + "<li>Flesch Ease    = 102  </li>\n";
  var ansFish1 = "<li>Coleman-Liau   = -0.07 (>-5) </li>\n"
               + "<li>Flesch-Kincaid = 0.54 (>-5) </li>\n"
               + "<li>Flesch Ease    = 102 (<200) </li>\n";
  var ansGett0 = "<li>Coleman-Liau   = 7  </li>\n"
               + "<li>Flesch-Kincaid = 5.31  </li>\n"
               + "<li>Flesch Ease    = 72  </li>\n";
  var ansGett1 = "<li>Coleman-Liau   = 7 (>5) </li>\n"
               + "<li>Flesch-Kincaid = 5.31 (>5) </li>\n"
               + "<li>Flesch Ease    = 72 (<100) </li>\n";
  assert_( ""       , 'checkReadability( fish                            )' );
  assert_( ""       , 'checkReadability( fish              , -5, -5, 200 )' );
  assert_( ansFish0 , 'checkReadability( fish + fish + fish              )' );
  assert_( ansFish1 , 'checkReadability( fish + fish + fish, -5, -5, 200 )' );
  assert_( ""       , 'checkReadability( fish + fish + fish,  5,  5,  50 )' );
  assert_( ansGett0 , 'checkReadability( getty                           )' );
  assert_( ansGett1 , 'checkReadability( getty             ,  5,  5, 100 )' );
  assert_( ""       , 'checkReadability( getty             , 10, 10,  50 )' );
}
function testCheckReadability_Null_() {
  assert_( "", 'checkReadability( "", 0, 0, 1000 )' );
}

function testClearHeading_() {
  assert_( 0, 'clearHeading( body, "Heading4" )' );
  body.insertParagraph(0,"XYZ").setHeading( DocumentApp.ParagraphHeading.HEADING4 );
  body.insertParagraph(0,"XYZ").setHeading( DocumentApp.ParagraphHeading.HEADING4 );
  assert_( 2, 'clearHeading( body, "Heading4" )' );
  assert_( 0, 'clearHeading( body, "Heading4" )' );
}
function testClearHeading_Null_() {
  assert_( 0, 'clearHeading( null, ""   )' );
  assert_( 0, 'clearHeading( body, null )' );
}

function testCountHeading_() {
  assert_( 0, 'countHeading( body, "Heading4" )' );
  body.insertParagraph(0,"XYZ").setHeading( DocumentApp.ParagraphHeading.HEADING4 );
  body.insertParagraph(0,"XYZ").setHeading( DocumentApp.ParagraphHeading.HEADING4 );
  assert_( 2, 'countHeading( body, "Heading4" )' );
}
function testCountHeading_Null_() {
  assert_( 0, 'countHeading( null, ""   )' );
  assert_( 0, 'countHeading( body, null )' );
}

function testCountLines_() {
  var text = "abcdefghijabcdefghijabcdefghijabcdefghijabcdefghijabcdefghijabcdefghijabcdefghij";
  assert_( 1, 'countLines( "' + text               + '")' );
  assert_( 2, 'countLines( "' + text + text        + '")' );
  assert_( 3, 'countLines( "' + text + text + text + '")' );
}
function testCountLines_Escape_() {
  assert_( 2, 'countLines( "\\r" )' );
  assert_( 5, 'countLines( "\\n\\r\\r\\r\\n" )' ); // final \r\n counts as 1, not 2
}
function testCountLines_Null_() {
  assert_( 0, 'countLines( null )' );
  assert_( 1, 'countLines( ""   )' );
}

function testCountPageBreak_() {
  assert_( 0, 'countPageBreak( null )' );
  assert_( 0, 'countPageBreak( body )' );
  body.insertPageBreak(0);
  assert_( 1, 'countPageBreak( body )' );
  body.insertPageBreak(0);
  body.insertPageBreak(0);
  assert_( 3, 'countPageBreak( body )' );
}

function testCountText_() {
  assert_( 0, 'countText( body.getText(), "XYZ" )' );
  body.insertParagraph(0, "XYZXYZ");
  assert_( 2, 'countText( body.getText(), "XYZ" )' );
  assert_( 0, 'countText( body.getText(), "WXY" )' );
}
function testCountText_Null_() {
  assert_( 0, 'countText( null,           "XYZ" )' );
  assert_( 0, 'countText( body.getText(), null  )' );
}

function testGetNewName_() {
  assert_( 'XYZ (Sample)' , 'getNewName("XYZ"          , props, "Sample"  )' );
  assert_( 'XYZ (Sample)' , 'getNewName("XYZ (Author)" , props, "Sample"  )' );
  assert_( 'XYZ (Sample)' , 'getNewName("XYZ  (Author)", props, "Sample"  )' );
  assert_( 'XYZ (Student)', 'getNewName("XYZ"          , props, "Student" )' );
  assert_( 'XYZ (Student)', 'getNewName("XYZ (Author)" , props, "Student" )' );
  assert_( 'XYZ (Student)', 'getNewName("XYZ  (Author)", props, "Student" )' );
  assert_( 'XYZ (Teacher)', 'getNewName("XYZ"          , props, "Teacher" )' );
  assert_( 'XYZ (Teacher)', 'getNewName("XYZ (Author)" , props, "Teacher" )' );
  assert_( 'XYZ (Teacher)', 'getNewName("XYZ  (Author)", props, "Teacher" )' );
}

function testGetReadabilityColemanLiau_() {
  assert_( -45.4 , 'getReadability("ColemanLiau"  , null            )', 0.1 );
  assert_( -45.4 , 'getReadability("ColemanLiau"  , ""              )', 0.1 );
  assert_(  -3.4 , 'getReadability("ColemanLiau"  , "' + fish  + '" )', 0.1 );
  assert_(   7.0 , 'getReadability("ColemanLiau"  , "' + getty + '" )', 0.1 );
}

function testGetReadabilityFleschEase_() {
  assert_(   206 , 'getReadability("FleschEase"   , null            )' );
  assert_(   206 , 'getReadability("FleschEase"   , ""              )' );
  assert_(   111 , 'getReadability("FleschEase"   , "' + fish  + '" )' );
  assert_(    72 , 'getReadability("FleschEase"   , "' + getty + '" )' );
}

function testGetReadabilityFleschKincaid_() {
  assert_( -15.2 , 'getReadability("FleschKincaid", null            )', 0.1 );
  assert_( -15.2 , 'getReadability("FleschKincaid", ""              )', 0.1 );
  assert_(  -1.1 , 'getReadability("FleschKincaid", "' + fish  + '" )', 0.1 );
  assert_(   5.3 , 'getReadability("FleschKincaid", "' + getty + '" )', 0.1 );
}

function testGetAlphanumerics_() {
  assert_( ""                                    , 'getAlphanumerics( null                                       )' );
  assert_( ""                                    , 'getAlphanumerics( ""                                         )' );
  assert_( ""                                    , 'getAlphanumerics( "-=!@#$%^&*()+[]\;,./{}|:<>?"              )' );
  assert_( "q1w2e3r4t5y6u7i8o9p0QWERTYUIOP"      , 'getAlphanumerics( "q1w2e3r4t5y6u7i8o9p0Q!W@E#R$T%Y^U&I*O(P)" )' );
}
function testGetWords_() {
  assert_( [ "" ]                                , 'getWords        ( null                      )' );
  assert_( [ "" ]                                , 'getWords        ( ""                        )' );
  assert_( [ "the", "quick", "brown", "dog", "" ], 'getWords        ( "the?quick !brown. dog; " )' );
}
function testGetSentences_() {
  assert_( [ "" ]                                , 'getSentences    ( null                      )' );
  assert_( [ "" ]                                , 'getSentences    ( ""                        )' );
  assert_( [ "the", "quick", "brown", "dog; " ]  , 'getSentences    ( "the?quick !brown. dog; " )' );
}
function testGetTitle_() {
  assert_( "",    'getTitle( null )' );
  assert_( "",    'getTitle( body )' );
  body.insertParagraph(0, "XYZ").setHeading( DocumentApp.ParagraphHeading.TITLE );
  assert_( "XYZ", 'getTitle( body )' );
}

function testHasText_Body_() {
  assert_( false, 'hasText( body, "XYZ" )' );
  body.insertParagraph(0, "XYZ");
  assert_(  true, 'hasText( body, "XYZ" )' );
  assert_( false, 'hasText( body, "WXY" )' );
}
function testHasText_Child_() {
  assert_( false, 'hasText( body.getChild(0), "XYZ" )' );
  body.insertParagraph(0, "XYZ");
  assert_(  true, 'hasText( body.getChild(0), "XYZ" )' );
  assert_( false, 'hasText( body.getChild(0), "WXY" )' );
}
function testHasText_Null_() {
  assert_( false, 'hasText( null, "XYZ" )' );
  assert_( false, 'hasText( body, null  )' );
}

function testRemoveAround_None_() {
  assert_(  0, 'removeAround( body, "START", "STOP" )' );
  assert_(  1, 'body.getNumChildren()' );
}

function testRemoveAround_None_() {
  assert_(  5, 'appendParagraphs_( "BEF", "MIDA", "MIDB", "MIDC", "AFT" )' );
  assert_(  6, 'body.getNumChildren()' );
  assert_(     '\nBEF\nMIDA\nMIDB\nMIDC\nAFT', 'body.getText()' );
  assert_(  0, 'removeAround( body, "START", "STOP" )' );
  assert_(  6, 'body.getNumChildren()' );
  assert_(     '\nBEF\nMIDA\nMIDB\nMIDC\nAFT', 'body.getText()' );
  assert_(  0, 'removeAround( body, "START", "STOP" )' );
  assert_(  6, 'body.getNumChildren()' );
  assert_(     '\nBEF\nMIDA\nMIDB\nMIDC\nAFT', 'body.getText()' );
}
function testRemoveAround_START_() {
  assert_(  3, 'appendParagraphs_( "BEF", "START", "AFT" )' );
  assert_(  4, 'body.getNumChildren()' );
  assert_(     '\nBEF\nSTART\nAFT'     , 'body.getText()' );
  assert_(  3, 'removeAround( body, "START", "STOP" )' );
  assert_(  1, 'body.getNumChildren()' );
  assert_(     'AFT'                   , 'body.getText()' );
}
function testRemoveAround_STOP_() {
  assert_(  3, 'appendParagraphs_( "BEF", "STOP", "AFT" )' );
  assert_(  4, 'body.getNumChildren()' );
  assert_(     '\nBEF\nSTOP\nAFT'      , 'body.getText()' );
  assert_(  2, 'removeAround( body, "START", "STOP" )' );
  assert_(  3, 'body.getNumChildren()' );
  assert_(     '\nBEF\n'               , 'body.getText()' );
}
function testRemoveAround_Simple_() {
  assert_(  9, 'appendParagraphs_( "BEF", "START", "MIDA", "STOP", "MIDB", "START", "MIDC", "STOP", "AFT" )' );
  assert_( 10, 'body.getNumChildren()' );
  assert_(     '\nBEF\nSTART\nMIDA\nSTOP\nMIDB\nSTART\nMIDC\nSTOP\nAFT', 'body.getText()' );
  assert_(  8, 'removeAround( body, "START", "STOP" )' );
  assert_(  3, 'body.getNumChildren()' );
  assert_(     'MIDA\nMIDC\n', 'body.getText()' );
  assert_(  0, 'removeAround( body, "START", "STOP" )' );
  assert_(  3, 'body.getNumChildren()' );
  assert_(     'MIDA\nMIDC\n', 'body.getText()' );
}
function testRemoveAround_Repeated_() {
  assert_(  7, 'appendParagraphs_( "BEF", "START", "START", "MIDA", "STOP", "STOP", "AFT" )' );
  assert_(  8, 'body.getNumChildren()' );
  assert_(     '\nBEF\nSTART\nSTART\nMIDA\nSTOP\nSTOP\nAFT', 'body.getText()' );
  assert_(  7, 'removeAround( body, "START", "STOP" )' );
  assert_(  2, 'body.getNumChildren()' );
  assert_(     'MIDA\n', 'body.getText()' );
  assert_(  0, 'removeAround( body, "START", "STOP" )' );
  assert_(  2, 'body.getNumChildren()' );
  assert_(     'MIDA\n', 'body.getText()' );
}
function testRemoveAround_Nested_() {
  assert_(  9, 'appendParagraphs_( "BEF", "START", "MIDA", "START", "MIDB", "STOP", "MIDC", "STOP", "AFT" )' );
  assert_( 10, 'body.getNumChildren()' );
  assert_(     '\nBEF\nSTART\nMIDA\nSTART\nMIDB\nSTOP\nMIDC\nSTOP\nAFT', 'body.getText()' );
  assert_(  9, 'removeAround( body, "START", "STOP" )' );
  assert_(  2, 'body.getNumChildren()' );
  assert_(     'MIDB\n', 'body.getText()' );
  assert_(  0, 'removeAround( body, "START", "STOP" )' );
  assert_(  2, 'body.getNumChildren()' );
  assert_(     'MIDB\n', 'body.getText()' );
}

function testRemoveFooter_() {
  assert_( false, 'removeFooter( doc, true )');
  doc.addFooter();
  assert_(  true, 'removeFooter( doc, true )');
  assert_( false, 'removeFooter( doc, true )');
}

function testRemoveHeader_() {
  assert_( false, 'removeHeader( doc, true )');
  doc.addHeader();
  assert_(  true, 'removeHeader( doc, true )');
  assert_( false, 'removeHeader( doc, true )');  
}

function testRepeat_() {
  assert_(    "", 'repeat( "X", -1 )' );
  assert_(    "", 'repeat( "X",  0 )' );
  assert_(   "X", 'repeat( "X",  1 )' );
  assert_( "XXX", 'repeat( "X",  3 )' );  
}

function testReplaceImageTag_() {
  assert_(    0, 'replaceImageTag( body, "TAG", "http://kussmaul.org/clif.jpg", "http://kussmaul.org" )' );
  body.appendParagraph("BEF-TAG-AFT");
  assert_(    2, 'body.getNumChildren()'                                                                 );
  assert_(       '\nBEF-TAG-AFT', 'body.getText()' );
  assert_(    1, 'replaceImageTag( body, "TAG", "http://kussmaul.org/clif.jpg", "http://kussmaul.org" )' );
  assert_(    2, 'body.getNumChildren()'                                                                 );
  assert_( true, 'null != body.findElement(DocumentApp.ElementType.INLINE_IMAGE)'                        );
  assert_(       '\nBEF--AFT', 'body.getText()' );
  assert_(    0, 'replaceImageTag( body, "TAG", "http://kussmaul.org/clif.jpg", "http://kussmaul.org" )' );
}
function testReplaceImageTag_Null_() {
  body.appendParagraph("BEF-TAG-AFT");
  assert_(    2, 'body.getNumChildren()'                                                                 );
  assert_(       '\nBEF-TAG-AFT', 'body.getText()' );
  assert_(    0, 'replaceImageTag( null, "TAG", "http://kussmaul.org/clif.jpg", "http://kussmaul.org" )' );
  assert_(    0, 'replaceImageTag( body,  null, "http://kussmaul.org/clif.jpg", "http://kussmaul.org" )' );
  assert_(    0, 'replaceImageTag( body, "TAG", null                          , "http://kussmaul.org" )' );
}
function testReplaceImageTag_Two_() {
  assert_(    0, 'replaceImageTag( body, "TAG", "http://kussmaul.org/clif.jpg", "http://kussmaul.org" )' );
  body.appendParagraph("BEF-TAG-TAG-AFT");
  assert_(    2, 'body.getNumChildren()'                                                                 );
  assert_(       '\nBEF-TAG-TAG-AFT', 'body.getText()' );
  assert_(    2, 'replaceImageTag( body, "TAG", "http://kussmaul.org/clif.jpg", "http://kussmaul.org" )' );
  assert_(    2, 'body.getNumChildren()'                                                                 );
  assert_( true, 'null != body.findElement(DocumentApp.ElementType.INLINE_IMAGE)'                        );
  assert_(       '\nBEF---AFT', 'body.getText()' );
  assert_(    0, 'replaceImageTag( body, "TAG", "http://kussmaul.org/clif.jpg", "http://kussmaul.org" )' );
}

function testReplaceTextTag_() {
  assert_(    0, 'replaceTextTag( body, "TAG", "Clif Kussmaul", "http://kussmaul.org" )' );
  body.appendParagraph("BEF-TAG-AFT");
  assert_(    2, 'body.getNumChildren()'                                                 );
  assert_(       '\nBEF-TAG-AFT'                              , 'body.getText()' );
  assert_(    1, 'replaceTextTag( body, "TAG", "Clif Kussmaul", "http://kussmaul.org" )' );
  assert_(    2, 'body.getNumChildren()'                                                 );
  assert_( true, 'hasText( body.getChild(1), "Clif Kussmaul" )'                          );
  assert_(       '\nBEF-Clif Kussmaul-AFT'                    , 'body.getText()' );
  assert_(    0, 'replaceTextTag( body, "TAG", "Clif Kussmaul", "http://kussmaul.org" )' );
}
function testReplaceTextTag_Null_() {
  body.appendParagraph("BEF-TAG-AFT");
  assert_(    0, 'replaceTextTag( null, "TAG", "Clif Kussmaul", "http://kussmaul.org" )' );
  assert_(    0, 'replaceTextTag( body,  null, "Clif Kussmaul", "http://kussmaul.org" )' );
  assert_(    0, 'replaceTextTag( body, "TAG", null           , "http://kussmaul.org" )' );
  assert_(    1, 'replaceTextTag( body, "TAG", ""             , "http://kussmaul.org" )' );
}
// replaces 2 tags, both in same section
function testReplaceTextTag_Two_() {
  assert_(    0, 'replaceTextTag( body, "TAG", "Clif Kussmaul", "http://kussmaul.org" )' );
  body.appendParagraph("BEF-TAGTAG-AFT");
  assert_(    2, 'body.getNumChildren()'                                                 );
  assert_(       '\nBEF-TAGTAG-AFT'                           , 'body.getText()' );
  assert_(    1, 'replaceTextTag( body, "TAG", "Clif Kussmaul", "http://kussmaul.org" )' );
  assert_(    2, 'body.getNumChildren()'                                                 );
  assert_( true, 'hasText( body.getChild(1), "Clif Kussmaul" )'                          );
  assert_(       '\nBEF-Clif KussmaulClif Kussmaul-AFT'       , 'body.getText()' );
  assert_(    0, 'replaceTextTag( body, "TAG", "Clif Kussmaul", "http://kussmaul.org" )' );
}

function testSetFooter_() {
  assert_(  0,        'setFooter(  doc, "FT" )' );
  assert_(  "FT", 'doc.getFooter().getText()'   ); 
  assert_(  0,        'setFooter(  doc, null )' );
  if ( null != doc.getFooter() ) doc.getFooter().removeFromParent();
}
function testSetFooter_Null_() {
  assert_( false, 'setFooter( null, "FT" )' );
  if ( null != doc.getFooter() ) doc.getFooter().removeFromParent();
  assert_( false, 'setFooter(  doc, null )' );
}

function testSetHeader_() {
  assert_( 0   ,     'setHeader(  doc, "HD" )' );
  assert_( "HD", 'doc.getHeader().getText()'  ); 
  assert_( 0,        'setHeader(  doc, null )' );
  if ( null != doc.getHeader() ) doc.getHeader().removeFromParent();
}
function testSetHeader_Null_() {
  assert_( false, 'setHeader( null, "HD" )' );
  if ( null != doc.getHeader() ) doc.getHeader().removeFromParent();
  assert_( false, 'setHeader(  doc, null )' );
}
