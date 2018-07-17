(function () {
  Office.initialize = function (reason) {
    display('loaded');
    $(document).ready(function () {
      if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
        //never gets in here
        display("Requirement supported.");
      }
      else {
        display("This add-in requires Word 2016 or greater.");
      }
      display('document ready');
      var words = window.localStorage.getItem('words') || 'word1, word2';
      document.getElementById('words').innerText = words;
      display('words: ' + words);
      if (words) {
        window.words = getArr(words);
      }
      // Do something that is only available via the new APIs
      $('#findWords').click(findWords);
      $('#highlightWords').click(highlightWords);
      $('#change').click(change);
      $('#test').click(test);
      $(document).on('click', '.selectWord', select);
      $('#words').on('change', function (e) {
        var value = e.target.value;
        display(value);
        window.words = getArr(value);
        window.localStorage.setItem('words', value);
      });
      $('#findWordsExpand').click(function () {
        var toggle = document.getElementById('findWordsExpand');
        var result = document.getElementById('findResult');
        if (toggle.classList.contains('rotate')) {
          toggle.classList.remove('rotate');
          result.classList.remove('hide');
        }
        else {
          toggle.classList.add('rotate');
          result.classList.add('hide');
        }
      });
      $('#highlightWordsExpand').click(function () {
        var toggle = document.getElementById('highlightWordsExpand');
        var result = document.getElementById('highlightResult');
        if (toggle.classList.contains('rotate')) {
          toggle.classList.remove('rotate');
          result.classList.remove('hide');
        }
        else {
          toggle.classList.add('rotate');
          result.classList.add('hide');
        }
      });
    })
  };

  function display(s) {
    document.getElementById('wut').innerHTML += s + '</br>';
  }
  console.log = display;
  console.err = display;

  function highlightWords() {
    Word.run(function (context) {
      // Create a proxy object for the document.
      var thisDocument = context.document;
      var rangeList = [];
      var o = 0;
      window.words.forEach(function (word) {
        var ranges = thisDocument.body.search(word, { matchCase: false });
        context.load(ranges, 'font');
        context.load(ranges, 'text');
        rangeList.push(ranges);
      });
      return context.sync().then(function () {
        rangeList.forEach(function (ranges) {
          for (var i = 0; i < ranges.items.length; i++) {
            ranges.items[i].font.color = 'purple';
            ranges.items[i].font.highlightColor = '#FFFF00'; //Yellow
            ranges.items[i].font.bold = true;
            o++;
            document.getElementById("highlightCount").innerText = o;
          }
        })
        console.log('highlight words done');
        // range.font.color = '#FF0000';
        // console.log(html.value);
      });
    })
    .catch(display);
  }

  function select(e) {
    display('start select');
try {
display(e.currentTarget);
    var word = $(e.currentTarget).attr('word');
    var index = parseInt($(e.currentTarget).attr('index'));
    display('ohhh ' + word + ' ' + index);
    Word.run(function (context) {
      // Create a proxy object for the document.
      var thisDocument = context.document;
      var rangeList = [];
      var o = 0;
      var ranges = thisDocument.body.search(word, { matchCase: false });
      context.load(ranges, 'font');
      rangeList.push(ranges);
      return context.sync().then(function () {
        ranges.items[index].select();
        console.log('select words done');
        // range.font.color = '#FF0000';
        // console.log(html.value);
      });
    })
    .catch(display);
} catch(e) {display(e)};
  }

  function findWords() {
    Word.run(function (context) {
      var doc = context.document;
      var body = doc.body;
      context.load(body, 'text');
      var xMap = {};
      return context.sync().then(function () {
        var text = body.text;
        var resultHtml = '';
        var o = 0;
        var wordTemp = '(';
        window.words.forEach(function (word) {
          wordTemp += word + '|';
        });
        wordTemp = wordTemp.substr(0, wordTemp.length - 1) + ')';
        display(wordTemp)
        var reg = new RegExp(wordTemp, 'gi');
        var result;
        while ((result = reg.exec(text)) != null) {
          var word = result[0];
          xMap[word] = xMap[word] || 0;
          xMap[word]++;
          o++;
          resultHtml += (
            '<span class="selectWord" word="' + word + '" index="' + (xMap[word] - 1) + ')">'
            + o
            + ':&nbsp;&nbsp;&nbsp;&nbsp;'
            + text.substr(reg.lastIndex - word.length - 10, 10)
            + '<b>' + text.substr(reg.lastIndex - word.length, word.length) + '</b>'
            + text.substr(reg.lastIndex, 10)
            + '</span>'
            + '<br/>'
          );
          document.getElementById("findCount").innerText = o;
        }
        document.getElementById("findCount").innerText = o;
        document.getElementById("findResult").innerHTML = resultHtml;
        console.log('find words done');
      });
    })
    .catch(display);
  }

  function test() {
    Word.run(function (context) {
      // Create a proxy object for the document.
      var thisDocument = context.document;

      var ranges = thisDocument.body.search('your', { matchCase: false });
      context.load(ranges, 'font');
      return context.sync().then(function () {
        ranges.getFirst().select();
        for (var i = 0; i < ranges.items.length; i++) {
          ranges.items[i].font.color = 'purple';
          ranges.items[i].font.highlightColor = '#FFFF00'; //Yellow
          ranges.items[i].font.bold = true;
        }
        // range.font.color = '#FF0000';
        // console.log(html.value);
      });
    })
    .catch(function (error) {
      console.log('Error: ' + JSON.stringify(error));
      if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
      }
    });
  }

  function change() {
    Word.run(function (context) {
      // Create a proxy object for the document.
      var thisDocument = context.document;

      // Queue a command to get the current selection.
      // Create a proxy range object for the selection.
      var range = thisDocument.getSelection();
      // context.load(range, 'font');
      var font = range.font;
      context.load(font, 'bold');

      // Queue a command to replace the selected text.

      // Synchronize the document state by executing the queued commands,
      // and return a promise to indicate task completion.
      return context.sync().then(function () {
        font.bold = !font.bold;
        font.color = '#FF0000';
        font.underline = 'Double';
        // range.insertText(font.color.bold(), Word.InsertLocation.replace);
        console.log('changed');
      });
    })
    .catch(function (error) {
      console.log('Error: ' + JSON.stringify(error));
      if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
      }
    });
  }

  function getArr(csv) {
    try {
      return csv.split(',').map(function (x) {
        return x.trim();
      }).filter(function (x) {
        return x.length > 0;
      });
    }
    catch (e) {
      display(e);
      return [];
    }
  }
})();
