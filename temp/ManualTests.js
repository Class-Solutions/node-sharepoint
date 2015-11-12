'use strict';
var request = require('request');
var samlRequestTemplate = '<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope" xmlns:a="http://www.w3.org/2005/08/addressing" xmlns:u="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd"> <s:Header> <a:Action s:mustUnderstand="1">http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</a:Action> <a:ReplyTo> <a:Address>http://www.w3.org/2005/08/addressing/anonymous</a:Address> </a:ReplyTo> <a:To s:mustUnderstand="1">https://login.microsoftonline.com/extSTS.srf</a:To> <o:Security xmlns:o="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd" s:mustUnderstand="1"> <o:UsernameToken u:Id="uuid-1c603278-8629-4e34-ba09-ecd975c9a054-1"> <o:Username>lucas.dev@class-solutions.com.br</o:Username> <o:Password Type="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-username-token-profile-1.0#PasswordText"> Class@2015! </o:Password> </o:UsernameToken> </o:Security> </s:Header> <s:Body> <t:RequestSecurityToken xmlns:t="http://schemas.xmlsoap.org/ws/2005/02/trust"> <wsp:AppliesTo xmlns:wsp="http://schemas.xmlsoap.org/ws/2004/09/policy"> <a:EndpointReference> <a:Address>https://classsolutions.sharepoint.com/_forms/default.aspx?wa=wsignin1.0</a:Address> </a:EndpointReference> </wsp:AppliesTo> <t:KeyType>http://schemas.xmlsoap.org/ws/2005/05/identity/NoProofKey</t:KeyType> <t:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</t:RequestType> <t:TokenType>urn:oasis:names:tc:SAML:1.0:assertion</t:TokenType> </t:RequestSecurityToken> </s:Body></s:Envelope>';
process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";
module.exports.connect = function (cb) {
    request.post({
        uri: 'https://login.microsoftonline.com/extSTS.srf',
        proxy: "http://127.0.0.1:8888",
        body: samlRequestTemplate
    }, function (err, response, body) {
        console.log(body);
        console.log(response);
        console.log(err);
        cb();
    });
};
