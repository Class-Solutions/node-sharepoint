var assert = require('assert');

// it('should connect',function(done){
// 	this.timeout(50000);
// 	var SP = require('../temp/ManualTests.js');
// 	SP.connect(done);
// });

describe('Start Connection Test', function () {
	var SP;
	before(function () {
		SP = require('../temp/sharepoint.js');
	});
	describe('Testing Module "SP"', function () {
		it('Require should not be null', function () {
			assert.notEqual(SP, null);
		});
	});
	describe('Constructor for SP.RestService', function () {
		var client;
		before(function () {
			client = new SP.RestService('http://classsolutions.sharepoint.com/sites/DevAraujo');
		});
		it('should not be null', function () {
			assert.notEqual(client, null);
		});
		it('the client should connect', function (done) {
			this.timeout(30000);
			assert.doesNotThrow(
				function () {
					client.SignIn('lucas.dev@class-solutions.com.br', 'Class@2015!', function (err, data) {
						assert.equal(err, null);
						done();
					});
				},
				function (error) {
					return done(error);
				});
		});
		// it('should return the $metadata from SharePoint',function(done){
		// 	this.timeout(30000);
		// 	var c = client.Metadata(function(err,data){
		// 		console.log(err);
		// 		console.log(data);
		// 		done();
		// 	});
		// });
	});
});
