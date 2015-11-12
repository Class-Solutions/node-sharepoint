/// <reference path='../typings/node/node.d.ts' />

'use strict';

process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0"
var _ = require('lodash');
var fs = require('fs');
var qs = require('querystring');
var xml2js = require('xml2js');
var http = require('http');
var https = require('https');
var request = require('request');
var urlparse = require('url').parse;
//var samlRequestTemplate = fs.readFileSync(__dirname + '/SAML.xml', 'utf8');
var samlRequestTemplate = '<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope" xmlns:a="http://www.w3.org/2005/08/addressing" xmlns:u="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd"><s:Header><a:Action s:mustUnderstand="1">http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</a:Action><a:ReplyTo><a:Address>http://www.w3.org/2005/08/addressing/anonymous</a:Address></a:ReplyTo><a:To s:mustUnderstand="1">https://login.microsoftonline.com/extSTS.srf</a:To><o:Security s:mustUnderstand="1" xmlns:o="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"><o:UsernameToken><o:Username>[username]</o:Username><o:Password>[password]</o:Password></o:UsernameToken></o:Security></s:Header><s:Body><t:RequestSecurityToken xmlns:t="http://schemas.xmlsoap.org/ws/2005/02/trust"><wsp:AppliesTo xmlns:wsp="http://schemas.xmlsoap.org/ws/2004/09/policy"><a:EndpointReference><a:Address>[endpoint]</a:Address></a:EndpointReference></wsp:AppliesTo><t:KeyType>http://schemas.xmlsoap.org/ws/2005/05/identity/NoProofKey</t:KeyType><t:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</t:RequestType><t:TokenType>urn:oasis:names:tc:SAML:1.0:assertion</t:TokenType></t:RequestSecurityToken></s:Body></s:Envelope>';


/**SharePoint Connection Module */
module SP {
    /**SharePoint Connection Metadata */
    class MetadataOptions {
        public List: string;
        public Accept: string;
        public Method: string;
        
        /**Instantiates an Object with Connection Metadata  */
        public constructor(
            /**List Name */
            list: string,
            /**Accepted return Content-Type */
            accept: string, 
            /**The Verb Method for the Request */
            method: string) {
            this.List = list;
            this.Accept = accept;
            this.Method = method;
        }
    }

    class Cookie {
        public Name: string;
        public Value: string;

        public constructor(name: string, value: string) {
            this.Name = name;
            this.Value = value;
        }
    }
    
    /**REST Requests Service */
    export class RestService {
        private Url: any;
        private Host: string;
        private Path: string;
        private Protocol: string;

        private STS: {
            Host: string;
            Path: string;
        }
        private Login: string;
        private ODataEndPoint: string;
        private FedAuth: string;
        private rtFa: string;
        
        /**Instantiates a new Instance of the REST Requests Service */
        public constructor(
            /**The SharePoint Site URL */
            url: any) {

            this.Url = urlparse(url);
            this.Host = this.Url.host;
            this.Path = this.Url.path;
            this.Protocol = this.Url.protocol;
            
            // External Security Token Service for SPO
            this.STS = {
                Host: 'login.microsoftonline.com',
                Path: '/extSTS.srf'
            };

            // Form to submit SAML token
            this.Login = '/_forms/default.aspx?wa=wsignin1.0';

            // SharePoint Odata (REST) service
            this.ODataEndPoint = '/_vti_bin/ListData.svc/';
        }
        
        /**Requests the Sign in of the informed user */
        public SignIn(
            /**User Login Name */
            username: string, 
            /**User Password */
            password: string, 
            /**Callback Function for the Login Event */
            callback: (error: any, data: any) => any): void {
            var self = this;

            var options = {
                username: username,
                password: password,
                sts: self.STS,
                endpoint: self.Url.protocol + '//' + self.Url.hostname + self.Login
            };

            RestService._RequestToken(options, (err: any, data: any): void=> {

                if (err) {
                    callback(err, null);
                    return;
                }

                self.FedAuth = data.FedAuth;
                self.rtFa = data.rtFa;

                callback(null, data);
            });
        }
        
        /**Requests a SP.List Object */
        public GetList(
            /**Name of the required List */
            name: string): SP.List {
            var list = new SP.List(this, name);
            return list;
        }

        public Request(options: any, next: (data: any, error: any) => any): void {
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
            // Include If-Match header if etag is specified
            if (options.etag) {
                req_options.headers['If-Match'] = options.etag;
            };
            var protocol = (ssl ? https : http);
            var req = protocol.request(req_options, (res: any): void=> {
                var res_data = '';
                res.setEncoding('utf8');
                res.on('data', (chunk: any): void=> {
                    res_data += chunk;
                });
                res.on('end', (): void=> {
                    // if no callback is defined, we're done.
                    if (!next) return;

                    // if data of content-type application/json is return, parse into JS:
                    if (res_data && (res.headers['content-type'].indexOf('json') > 0)) {
                        res_data = JSON.parse(res_data).d
                    }

                    if (res_data) {
                        next(null, res_data)
                    }
                    else {
                        next(null, null);
                    }
                });
            });
            req.end(req_data);
        }

        public Metadata(next: (ev: Event) => any): void {
            var options = new MetadataOptions('$metadata', 'application/xml', 'GET');

            this.Request(options, next);
        }

        static _BuildSamlRequest(params: any): any {
            var saml = samlRequestTemplate;

            for (var key in params) {
                saml = saml.replace('[' + key + ']', params[key]);
            }
            //console.log(saml);
            return saml;
        }

        static _ParseXml(xml: string, callback: (ev: Event) => any): void {
            var parser = new xml2js.Parser({
                emptyTag: ''  // use empty string as value when tag empty
            });

            parser.on('end', (js: Event): void => {
                callback && callback(js);
            });

            parser.parseString(xml);
        }

        static _ParseCookie(txt: string): any {
            var properties = txt.split('; ');
            var cookie = new Cookie('', '');

            properties.forEach(function(property, index) {
                var idx = property.indexOf('='),
                    name = (idx > 0 ? property.substring(0, idx) : property),
                    value = (idx > 0 ? property.substring(idx + 1) : undefined);

                if (index === 0) {
                    cookie.Name = name;
                    cookie.Value = value;
                } else {
                    cookie.Name = value
                }

            })

            return cookie;
        }

        static _ParseCookies(txts: Array<string>): any {
            var cookies = new Array<string>();

            if (txts) {
                txts.forEach((txt: string): void => {
                    var cookie = this._ParseCookie(txt);
                    cookies.push(cookie)
                });
            }

            return cookies;
        }

        static _GetCookie(cookies: Array<string>, name: string): any {
            var cookie: any;
            var len = cookies.length;

            for (var i = 0; i < len; i++) {
                cookie = cookies[i]
                if (cookie.name == name) {
                    return cookie
                }
            }

            return undefined;
        }

        static _SubmitToken(params: any, callback: any): any {
            var token = params.token,
                url = urlparse(params.endpoint),
                ssl = (url.protocol == 'https:');

            var options = {
                method: 'POST',
                host: url.hostname,
                path: url.path,
                headers: {
                    // E accounts require a user agent string
                    'User-Agent': 'Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Win64; x64; Trident/5.0)'
                }
            };

            var protocol = (ssl ? https : http);

            var req = protocol.request(options, (res: any): void=> {

                var xml = '';
                res.setEncoding('utf8');
                res.on('data', (chunk: any): void=> {
                    xml += chunk;
                })

                res.on('end', function() {

                    var cookies = RestService._ParseCookies(res.headers['set-cookie'])

                    callback(null, {
                        FedAuth: RestService._GetCookie(cookies, 'FedAuth').value,
                        rtFa: RestService._GetCookie(cookies, 'rtFa').value
                    })
                })
            });

            req.end(token);
        }

        static _RequestToken(params: any, callback: (error: any, data: any) => any): void {
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
                proxy: 'http://127.0.0.1:8888', // For Fiddler Feedback TODO: Remove in Production
                body: samlRequest
            }, (err: any, response: any, body: any): void=> {
                var xml = '';
                console.log(err);
                console.log(response);
                console.log(body);
                //                 res.setEncoding('utf8');
                //                 res.on('data', (chunk: any): void=> {
                //                     xml += chunk;
                //                 })
                //                 res.on('end', (): void=> {
                //                     RestService._ParseXml(xml, (js: any): void=> {
                //                         // check for errors
                //                         var Fault = _.get(js, 'S:Envelope.S:Body[0].S:Fault[0]');
                //                         if (Fault) {
                //                             var error = _.get(Fault, 'S:Detail[0].psf:error[0].psf:internalerror[0]');
                //                             var errorMessage = _.get(error, 'psf:text[0]');
                //                             var errorCode = _.get(error, 'psf:code')
                //                             callback(errorCode + ' ' + errorMessage, null);
                //                             return;
                //                         } 
                // 
                //                         // extract token
                //                         var token = js['S:Envelope']['S:Body']['wst:RequestSecurityTokenResponse']['wst:RequestedSecurityToken']['wsse:BinarySecurityToken']['#'];
                // 
                //                         // Now we have the token, we need to submit it to SPO
                //                         RestService._SubmitToken({
                //                             token: token,
                //                             endpoint: params.endpoint
                //                         }, callback)
                //                     })
                //                 })
                callback('', '');
            });

            //             var req = https.request(options, (res: any): void=> {
            //                 var xml = '';
            // 
            //                 res.setEncoding('utf8');
            //                 res.on('data', (chunk: any): void=> {
            //                     xml += chunk;
            //                 })
            //                 res.on('end', (): void=> {
            //                     RestService._ParseXml(xml, (js: any): void=> {
            //                         // check for errors
            //                         var Fault = _.get(js, 'S:Envelope.S:Body[0].S:Fault[0]');
            //                         if (Fault) {
            //                             var error = _.get(Fault, 'S:Detail[0].psf:error[0].psf:internalerror[0]');
            //                             var errorMessage = _.get(error, 'psf:text[0]');
            //                             var errorCode = _.get(error, 'psf:code')
            //                             callback(errorCode + ' ' + errorMessage, null);
            //                             return;
            //                         } 
            // 
            //                         // extract token
            //                         var token = js['S:Envelope']['S:Body']['wst:RequestSecurityTokenResponse']['wst:RequestedSecurityToken']['wsse:BinarySecurityToken']['#'];
            // 
            //                         // Now we have the token, we need to submit it to SPO
            //                         RestService._SubmitToken({
            //                             token: token,
            //                             endpoint: params.endpoint
            //                         }, callback)
            //                     })
            //                 })
            //             });
            // 
            //             req.end(samlRequest);
        }
    }
    
    /**Auxiliar Methods to Deal with SharePoint Lists */
    export class List {
        /**Instantiates an auxiliar object to deal with the List */
        public constructor(
            /**Service Name */ //TODO: Improve Documentation
            service: RestService, 
            /**List Name */
            name: string) {
        }
    }
}

module.exports = SP;