'use strict';
// nodemon -w . --exec npm test

var util = require('util');
var fs = require('fs');

//var log = function() {
    //var args = Array.prototype.slice.call(arguments, 0);
    //return util.log(util.inspect.call(null, args.length === 1 ? args[0] : args, false, null, true));
//};


module.exports.parse = function(assert) {
    var plist = require('../index');
    var fixture = JSON.parse(fs.readFileSync(__dirname + '/fixtures/test.json'));
    var filename = __dirname + '/fixtures/test.xlsx';
    var xlsObject;
    // parse file
    xlsObject = plist.parse(filename);
    assert.deepEqual(JSON.parse(JSON.stringify(xlsObject)), fixture, "Parse file asynchronously");
    // parse buffer
    xlsObject = plist.parse(fs.readFileSync(filename));
    assert.deepEqual(JSON.parse(JSON.stringify(xlsObject)), fixture, "Parse file to buffer");
    assert.done();
};

module.exports.parseFileAsync = function(assert) {
    var fixture = JSON.parse(fs.readFileSync(__dirname + '/fixtures/test.json'));
    var filename = __dirname + '/fixtures/test.xlsx';
    var plist = require('../index');
    plist.parseFileAsync(filename, {}, function( xlsxObject ){
        assert.done();
        assert.deepEqual(JSON.parse(JSON.stringify(xlsxObject)), fixture, "Parse file with promise");
    })
};
