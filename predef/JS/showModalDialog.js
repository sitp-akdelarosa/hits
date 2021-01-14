///========================================
/// showModalDialog.js
///
/// Created by Haraguroicha 2014-10-06
///========================================
(function() {
  window._smdName = window._smdName || Math.round(Math.random() * 1000000000);
  window.spawn = window.spawn || function(gen) {
    function continuer(verb, arg) {
      var result;
      try {
        result = generator[verb](arg);
      } catch (err) {
        return Promise.reject(err);
      }
      if (result.done) {
        return result.value;
      } else {
        return Promise.resolve(result.value).then(onFulfilled, onRejected);
      }
    }
    var generator = gen();
    var onFulfilled = continuer.bind(continuer, 'next');
    var onRejected = continuer.bind(continuer, 'throw');
    return onFulfilled();
  };
  window.showModalDialog = window.showModalDialog || function(url, arg, opt) {
    url = url || '';                                         // URL of a dialog
    arg = arg || null ;                                      // arguments to a dialog
    opt = opt || 'dialogWidth: 300px; dialogHeight: 200px';  // options: dialogTop;dialogLeft;dialogWidth;dialogHeight or CSS styles
    opt = opt
      .replace(/dialog/gi, '')                               // remove all of dialog strings
      .replace(/ /g, '')                                     // remove all blank characters
      .replace(/:/g, '= ')                                   // replace all of ':' to '= '
      .replace(/,|;/g, ', ')                                 // replace all of ',' or ';' to ', '
      .replace(/width/gi, 'width')                           // replace all 'width' to lowercase
      .replace(/height/gi, 'height')                         // replace all 'height' to lowercase
      .replace(/(\d+)px/g, '$1');                            // remove all of 'px'
    console.log(opt);
    var caller = showModalDialog.caller.toString();
    var dialog = window.open(url, 'smd_dialog_' + window._smdName, opt, false);
    dialog.dialogArguments = arg;
    dialog.addEventListener('unload', function(e) {
      e.preventDefault();
    });
    // if using yield
    if (caller.indexOf('yield') >= 0) {
      return new Promise(function(resolve, reject) {
        dialog.addEventListener('unload', function() {
          var returnValue = dialog.returnValue;
          resolve(returnValue);
        });
      });
    }
    // if using eval
    var isNext = false;
    var nextStmts = caller
      .replace(/(window\.)?showModalDialog\([^)]+\)/g, 'showModalDialog(%%%%%%%)')
      .split('\n')
      .filter(function(stmt) {
        if (isNext || stmt.indexOf('showModalDialog(') >= 0)
          return isNext = true;
        return false;
      });
    var unloadEventHandler = function() {
      if (dialog.location.href == 'about:blank') {
        return setTimeout(function() { dialog.addEventListener('unload', unloadEventHandler); }, 250);
      }
      var returnValue = dialog.returnValue;
      nextStmts[0] = nextStmts[0].replace(/(window\.)?showModalDialog\(%%%%%%%\)/g, JSON.stringify(returnValue));
      eval('{\n' + nextStmts.join('\n'));
    };
    dialog.addEventListener('unload', unloadEventHandler);
    throw 'Execution stopped until showModalDialog is closed';
  };
})();