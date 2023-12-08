/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Start', 'replaceRequirements')
      .addToUi();
}

/**
 * Runs when the add-on is installed.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
  onOpen(e);
}


function replaceRequirements() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  Logger.log('test');

  var regex = /REQ-\d+/g;
  var matches = body.getText().match(regex).filter((value, index, self) => {
    return self.indexOf(value) === index;
  });

  var apiUrl = 'https://ww1.requirementyogi.cloud/nuitdelinfo/search';
  var response = UrlFetchApp.fetch(apiUrl);
  var json = response.getContentText();
  var data = JSON.parse(json);
  var results = data.results;
  var reqList = {};
  for (var i = 0; i < results.length; i++) {
    reqList[results[i].key] = results[i].canonicalUrl;
  }

  var table = body.appendTable();

  if (matches) {

    for (var i = 0; i < matches.length; i++) {
      var requirementKey = matches[i];

        var textElement = body.findText(requirementKey);
        if (textElement && reqList[requirementKey]) {
          var row = table.appendTableRow();
          row.appendTableCell(requirementKey);
          linkText(requirementKey, reqList[requirementKey])
        }

    }
  }
}

/**
 * Find all matches of target text in current document, and add a url
 * Link to them if they don't already have one. The baseUrl may
 * include a placeholder, %target%, to be replaced with the matched text.
 *
 * @param {String} target   The text or regex to search for.
 *                          See Body.findText() for details.
 * @param {String} baseUrl  The URL that should be set on matching text.
 */
function linkText(target,baseUrl) {
  var doc = DocumentApp.getActiveDocument();
  var bodyElement = DocumentApp.getActiveDocument().getBody();
  var searchResult = bodyElement.findText(target);

  while (searchResult !== null) {
    var thisElement = searchResult.getElement();
    var thisElementText = thisElement.asText();
    var matchString = thisElementText.getText()
          .substring(searchResult.getStartOffset(), searchResult.getEndOffsetInclusive()+1);

    // if found text does not have a link already, add one
    if (thisElementText.getLinkUrl(searchResult.getStartOffset()) == null) {
      var url = baseUrl.replace('%target%',matchString)
      thisElementText.setLinkUrl(searchResult.getStartOffset(), searchResult.getEndOffsetInclusive(), url);
    }

    // search for next match
    searchResult = bodyElement.findText(target, searchResult);
  }
}
