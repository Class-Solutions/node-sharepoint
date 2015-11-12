/// <reference path='../typings/node/node.d.ts' />
'use strict';
process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";
var _ = require('lodash');
var fs = require('fs');
var qs = require('querystring');
var xml2js = require('xml2js');
var http = require('http');
var https = require('https');
var request = require('request');
var urlparse = require('url').parse;
var samlRequestTemplate = '<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope" xmlns:a="http://www.w3.org/2005/08/addressing" xmlns:u="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd"><s:Header><a:Action s:mustUnderstand="1">http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</a:Action><a:ReplyTo><a:Address>http://www.w3.org/2005/08/addressing/anonymous</a:Address></a:ReplyTo><a:To s:mustUnderstand="1">https://login.microsoftonline.com/extSTS.srf</a:To><o:Security s:mustUnderstand="1" xmlns:o="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"><o:UsernameToken><o:Username>[username]</o:Username><o:Password>[password]</o:Password></o:UsernameToken></o:Security></s:Header><s:Body><t:RequestSecurityToken xmlns:t="http://schemas.xmlsoap.org/ws/2005/02/trust"><wsp:AppliesTo xmlns:wsp="http://schemas.xmlsoap.org/ws/2004/09/policy"><a:EndpointReference><a:Address>[endpoint]</a:Address></a:EndpointReference></wsp:AppliesTo><t:KeyType>http://schemas.xmlsoap.org/ws/2005/05/identity/NoProofKey</t:KeyType><t:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</t:RequestType><t:TokenType>urn:oasis:names:tc:SAML:1.0:assertion</t:TokenType></t:RequestSecurityToken></s:Body></s:Envelope>';
var SP;
(function (SP) {
    var MetadataOptions = (function () {
        function MetadataOptions(list, accept, method) {
            this.List = list;
            this.Accept = accept;
            this.Method = method;
        }
        return MetadataOptions;
    })();
    var Cookie = (function () {
        function Cookie(name, value) {
            this.Name = name;
            this.Value = value;
        }
        return Cookie;
    })();
    var RestService = (function () {
        function RestService(url) {
            this.Url = urlparse(url);
            this.Host = this.Url.host;
            this.Path = this.Url.path;
            this.Protocol = this.Url.protocol;
            this.STS = {
                Host: 'login.microsoftonline.com',
                Path: '/extSTS.srf'
            };
            this.Login = '/_forms/default.aspx?wa=wsignin1.0';
            this.ODataEndPoint = '/_vti_bin/ListData.svc/';
        }
        RestService.prototype.SignIn = function (username, password, callback) {
            var self = this;
            var options = {
                username: username,
                password: password,
                sts: self.STS,
                endpoint: self.Url.protocol + '//' + self.Url.hostname + self.Login
            };
            RestService._RequestToken(options, function (err, data) {
                if (err) {
                    callback(err, null);
                    return;
                }
                self.FedAuth = data.FedAuth;
                self.rtFa = data.rtFa;
                callback(null, data);
            });
        };
        RestService.prototype.GetList = function (name) {
            var list = new SP.List(this, name);
            return list;
        };
        RestService.prototype.Request = function (options, next) {
            var req_data = options.data || '';
            var list = options.list;
            var id = options.id;
            var query = options.query;
            var ssl = (this.Protocol == 'https:');
            var path = this.Path + this.ODataEndPoint + list +
                (id ? '(' + id + ')' : '') +
                (query ? '?' + qs.stringify(query) : '');
            var req_options = {
                method: options.method,
                host: this.Host,
                path: path,
                headers: {
                    'Accept': options.accept || 'application/json',
                    'Content-type': 'application/json',
                    'Cookie': 'FedAuth=' + this.FedAuth + '; rtFa=' + this.rtFa,
                    'Content-length': req_data.length,
                    'If-Match': ''
                }
            };
            if (options.etag) {
                req_options.headers['If-Match'] = options.etag;
            }
            ;
            var protocol = (ssl ? https : http);
            var req = protocol.request(req_options, function (res) {
                var res_data = '';
                res.setEncoding('utf8');
                res.on('data', function (chunk) {
                    res_data += chunk;
                });
                res.on('end', function () {
                    if (!next)
                        return;
                    if (res_data && (res.headers['content-type'].indexOf('json') > 0)) {
                        res_data = JSON.parse(res_data).d;
                    }
                    if (res_data) {
                        next(null, res_data);
                    }
                    else {
                        next(null, null);
                    }
                });
            });
            req.end(req_data);
        };
        RestService.prototype.Metadata = function (next) {
            var options = new MetadataOptions('$metadata', 'application/xml', 'GET');
            this.Request(options, next);
        };
        RestService._BuildSamlRequest = function (params) {
            var saml = samlRequestTemplate;
            for (var key in params) {
                saml = saml.replace('[' + key + ']', params[key]);
            }
            return saml;
        };
        RestService._ParseXml = function (xml, callback) {
            var parser = new xml2js.Parser({
                emptyTag: ''
            });
            parser.on('end', function (js) {
                callback && callback(js);
            });
            parser.parseString(xml);
        };
        RestService._ParseCookie = function (txt) {
            var properties = txt.split('; ');
            var cookie = new Cookie('', '');
            properties.forEach(function (property, index) {
                var idx = property.indexOf('='), name = (idx > 0 ? property.substring(0, idx) : property), value = (idx > 0 ? property.substring(idx + 1) : undefined);
                if (index === 0) {
                    cookie.Name = name;
                    cookie.Value = value;
                }
                else {
                    cookie.Name = value;
                }
            });
            return cookie;
        };
        RestService._ParseCookies = function (txts) {
            var _this = this;
            var cookies = new Array();
            if (txts) {
                txts.forEach(function (txt) {
                    var cookie = _this._ParseCookie(txt);
                    cookies.push(cookie);
                });
            }
            return cookies;
        };
        RestService._GetCookie = function (cookies, name) {
            var cookie;
            var len = cookies.length;
            for (var i = 0; i < len; i++) {
                cookie = cookies[i];
                if (cookie.name == name) {
                    return cookie;
                }
            }
            return undefined;
        };
        RestService._SubmitToken = function (params, callback) {
            var token = params.token, url = urlparse(params.endpoint), ssl = (url.protocol == 'https:');
            var options = {
                method: 'POST',
                host: url.hostname,
                path: url.path,
                headers: {
                    'User-Agent': 'Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Win64; x64; Trident/5.0)'
                }
            };
            var protocol = (ssl ? https : http);
            var req = protocol.request(options, function (res) {
                var xml = '';
                res.setEncoding('utf8');
                res.on('data', function (chunk) {
                    xml += chunk;
                });
                res.on('end', function () {
                    var cookies = RestService._ParseCookies(res.headers['set-cookie']);
                    callback(null, {
                        FedAuth: RestService._GetCookie(cookies, 'FedAuth').value,
                        rtFa: RestService._GetCookie(cookies, 'rtFa').value
                    });
                });
            });
            req.end(token);
        };
        RestService._RequestToken = function (params, callback) {
            var samlRequest = RestService._BuildSamlRequest({
                username: params.username,
                password: params.password,
                endpoint: params.endpoint
            });
            var options = {
                method: 'POST',
                host: params.sts.Host,
                path: params.sts.Path,
                headers: {
                    'Content-Length': Buffer.byteLength(samlRequest.length)
                }
            };
            var req = request.post({
                uri: 'https://login.microsoftonline.com/extSTS.srf',
                proxy: 'http://127.0.0.1:8888',
                body: samlRequest
            }, function (err, response, body) {
                var xml = '';
                console.log(err);
                console.log(response);
                console.log(body);
                callback('', '');
            });
        };
        return RestService;
    })();
    SP.RestService = RestService;
    var List = (function () {
        function List(service, name) {
        }
        return List;
    })();
    SP.List = List;
})(SP || (SP = {}));
module.exports = SP;
