function onOpen() {
  var ui = DocumentApp.getUi();
  setWordGoal(0,0);
  ui.createMenu('Counter Menu')
    .addItem('Progress', 'progress')
    .addItem('Edit word count goal', 'wordCountGoal')
    .addItem('Help', 'help')
    .addToUi();
}

function getWordGoal() {
  var properties = PropertiesService.getDocumentProperties();
  return {
    minGoal: parseInt(properties.getProperty('minGoal'), 10),
    maxGoal: parseInt(properties.getProperty('maxGoal'), 10)
  };
}

function setWordGoal(minGoal, maxGoal) {
  var properties = PropertiesService.getDocumentProperties();
  properties.setProperty('minGoal', String(minGoal));
  properties.setProperty('maxGoal', String(maxGoal));
}

function counter() {
  var text = DocumentApp.getActiveDocument().getBody().getText();
  var words = text.match(/\p{L}+(?:[\p{M}'-]*\p{L}+)*/gu) || [];
  return words.length;
}

function progress() {
  var ui = DocumentApp.getUi();
  var { minGoal, maxGoal } = getWordGoal();
  var minText
  var maxText
  while (minGoal === 0 && maxGoal === 0) {
    ui.alert("Please put a word minimum and/or word maximum.");
    wordCountGoal();
    var { minGoal, maxGoal } = getWordGoal();
  }
  words = counter();
  if (minGoal != 0) {
    minPercentage = Math.round(words / minGoal * 10000) / 100;
    minText = (words + "/" + minGoal + " words. " + minPercentage + "% to word minimum.");
  } else {
    minText = "No minimum defined";
  }
  if (maxGoal != 0) {
    maxPercentage = Math.round(words / maxGoal * 10000) / 100;
    maxText = (words + "/" + maxGoal + " words. " + maxPercentage + "% to word limit.");
  } else {
    maxText = "No maximum defined";
  }
  ui.alert(minText + "\n\n" + maxText);
}

function wordCountGoal() {
  var ui = DocumentApp.getUi();
  var { minGoal, maxGoal } = getWordGoal();
  var minResult = ui.prompt(
      'Edit word count goal', 
      'Enter the minimum word count:',
      ui.ButtonSet.OK_CANCEL);
  if (minResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  minGoal = parseInt(minResult.getResponseText().trim(), 10);
  if (isNaN(minGoal) || minGoal < 0) {
    ui.alert("Invalid minimum word count. Please try again.");
    return;
  }

  var maxResult = ui.prompt(
      'Edit word count goal', 
      'Enter the maximum word count:',
      ui.ButtonSet.OK_CANCEL);
  if (maxResult.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  maxGoal = parseInt(maxResult.getResponseText().trim(), 10);
  if (isNaN(maxGoal) || maxGoal < 0 || maxGoal <= minGoal) {
    ui.alert("Invalid maximum word count. Please try again.");
    return;
  }
  if (minGoal == 0 && maxGoal == 0) {
    return;
  }
  setWordGoal(minGoal, maxGoal);
  ui.alert("Word count goals changed successfully!");
}

function help() {
  var ui = DocumentApp.getUi();
  ui.alert(`
Docs Counter Menu
By WinterTech13

Helps track word counts!

To add word limits, click \"Edit word count goal\". Enter 0 to clear any previously defined limits.

To check progress towards limits, click \"Progress\".

Enjoy!`)
}
