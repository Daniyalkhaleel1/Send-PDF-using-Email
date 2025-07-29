var base64 = null;

function form_onsave(executionContext) {
    debugger;
    var formContext = executionContext.getFormContext();
    if (formContext.getAttribute("ccc_senttocustomer") !== null && formContext.getAttribute("ccc_senttocustomer") !== undefined) {
        var senttocustomer = formContext.getAttribute("ccc_senttocustomer").getValue();
        if (senttocustomer === true) {
            createAttachment();
        }
    }
}
function getReportingSession() {
    var reportName = "AkhuwatReport.rdl"; //set this to the report you are trying to download
    //var reportGuid = "cffc6524-db7c-ed11-81ad-000d3aaa0146"; //set this to the guid of the report you are trying to download 
    var reportGuid = "cffc6524-db7c-ed11-81ad-000d3aaa0146"; //set this to the guid of the report you are trying to download 
    var rptPathString = ""; //set this to the CRMF_Filtered parameter            
    var selectedIds = Xrm.Page.data.entity.getId();
    var pth = Xrm.Page.context.getClientUrl() + "/CRMReports/rsviewer/reportviewer.aspx";
    var retrieveEntityReq = new XMLHttpRequest();

    var strParameterXML = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'><entity name='aka_donation'><all-attributes/><filter type='and'><condition attribute='aka_donationid' operator='eq' value='" + Xrm.Page.data.entity.getId() + "' /></filter></entity></fetch>";

    retrieveEntityReq.open("POST", pth, false);

    retrieveEntityReq.setRequestHeader("Accept", "*/*");

    retrieveEntityReq.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

    retrieveEntityReq.send("id=%7B" + reportGuid + "%7D&uniquename=" + Xrm.Page.context.getOrgUniqueName() + "&iscustomreport=true&reportnameonsrs=&reportName=" + reportName + "&isScheduledReport=false&p:CRM_quote=" + strParameterXML);
    var x = retrieveEntityReq.responseText.lastIndexOf("ReportSession=");
    var y = retrieveEntityReq.responseText.lastIndexOf("ControlID=");
    // alert("x" + x + "y" + y);
    var ret = new Array();

    ret[0] = retrieveEntityReq.responseText.substr(x + 14, 24);
    ret[1] = retrieveEntityReq.responseText.substr(x + 10, 32);
    return ret;
}

function createEntity(ent, entName, upd) {
    var jsonEntity = JSON.stringify(ent);
    var createEntityReq = new XMLHttpRequest();
    var ODataPath = Xrm.Page.context.getClientUrl() + "/XRMServices/2011/OrganizationData.svc";
    createEntityReq.open("POST", ODataPath + "/" + entName + "Set" + upd, false);
    createEntityReq.setRequestHeader("Accept", "application/json");
    createEntityReq.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    createEntityReq.send(jsonEntity);
    var newEntity = JSON.parse(createEntityReq.responseText).d;

    return newEntity;
}

function createAttachment() {
    var params = getReportingSession();

    if (msieversion() >= 1) {
        encodePdf_IEOnly(params);
    } else {
        encodePdf(params);
    }
}

var StringMaker = function () {
    this.parts = [];
    this.length = 0;
    this.append = function (s) {
        this.parts.push(s);
        this.length += s.length;
    }
    this.prepend = function (s) {
        this.parts.unshift(s);
        this.length += s.length;
    }
    this.toString = function () {
        return this.parts.join('');
    }
}

var keyStr = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";

function encode64(input) {
    var output = new StringMaker();
    var chr1, chr2, chr3;
    var enc1, enc2, enc3, enc4;
    var i = 0;

    while (i < input.length) {
        chr1 = input[i++];
        chr2 = input[i++];
        chr3 = input[i++];

        enc1 = chr1 >> 2;
        enc2 = ((chr1 & 3) << 4) | (chr2 >> 4);
        enc3 = ((chr2 & 15) << 2) | (chr3 >> 6);
        enc4 = chr3 & 63;

        if (isNaN(chr2)) {
            enc3 = enc4 = 64;
        } else if (isNaN(chr3)) {
            enc4 = 64;
        }

        output.append(keyStr.charAt(enc1) + keyStr.charAt(enc2) + keyStr.charAt(enc3) + keyStr.charAt(enc4));
    }

    return output.toString();
}

function encodePdf_IEOnly(params) {
    var bdy = new Array();
    var retrieveEntityReq = new XMLHttpRequest();

    var pth = Xrm.Page.context.getClientUrl() + "/Reserved.ReportViewerWebControl.axd?ReportSession=" + params[0] + "&Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=" + params[1] + "&OpType=Export&FileName=Public&ContentDisposition=OnlyHtmlInline&Format=PDF";
    retrieveEntityReq.open("GET", pth, false);
    retrieveEntityReq.setRequestHeader("Accept", "*/*");

    retrieveEntityReq.send();
    bdy = new VBArray(retrieveEntityReq.responseBody).toArray(); // minimum IE9 required

    createNotesAttachment(encode64(bdy));
}

function encodePdf(params) {
    var xhr = new XMLHttpRequest();
    var pth = Xrm.Page.context.getClientUrl() + "/Reserved.ReportViewerWebControl.axd?ReportSession=" + params[0] + "&Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=" + params[1] + "&OpType=Export&FileName=Public&ContentDisposition=OnlyHtmlInline&Format=PDF";
    xhr.open('GET', pth, true);
    xhr.responseType = 'arraybuffer';

    xhr.onload = function (e) {
        if (this.status == 200) {
            var uInt8Array = new Uint8Array(this.response);
            base64 = encode64(uInt8Array);
            CreateEmail(base64);
            createNotesAttachment(base64);
        }
    };
    xhr.send();
}

function msieversion() {

    var ua = window.navigator.userAgent;
    var msie = ua.indexOf("MSIE ");

    if (msie > 0 || !!navigator.userAgent.match(/Trident.*rv\:11\./))      // If Internet Explorer, return version number
        return parseInt(ua.substring(msie + 5, ua.indexOf(".", msie)));
    else                 // If another browser, return 0
        return 0;
}

function createNotesAttachment(base64data) {

    var propertyName;
    var propertyAddress;
    var propertyCity;
    var estimateNumber;
    if (Xrm.Page.getAttribute("name") !== null && Xrm.Page.getAttribute("name") !== undefined) {
        estimateNumber = Xrm.Page.getAttribute("name").getValue();
        //alert(estimateNumber);
    }
    if (Xrm.Page.getAttribute("ccc_propertyid") !== null && Xrm.Page.getAttribute("ccc_propertyid") !== undefined) {
        var propertyNameRef = Xrm.Page.getAttribute("ccc_propertyid");
        if (propertyNameRef != null && propertyNameRef != undefined) {
            propertyId = propertyNameRef.getValue()[0].id.slice(1, -1);
            propertyAddress = propertyNameRef.getValue()[0].name;
            var object = getPropertyAddress(propertyId);
            propertyName = object[1];
            propertyCity = object[2];
        }
    }

    var post = Object();
    post.DocumentBody = base64data;
    post.Subject = DonationRecipt;
    post.FileName = "  AkhuwatReport + ".pdf";
    post.MimeType = "application/pdf";
    post.ObjectId = Object();
    post.ObjectId.LogicalName = Xrm.Page.data.entity.getEntityName();
    post.ObjectId.Id = Xrm.Page.data.entity.getId();
    createEntity(post, "Annotation", "");
}

function CreateEmail(base64) {

    var recordURL;
    var propertyName;
    var propertyAddress;
    var propertyCity;
    var estimateNumber;
    var serverURL = Xrm.Page.context.getClientUrl();
    var email = {};
    var qid = Xrm.Page.data.entity.getId().replace(/[{}]/g, "");
    var OwnerLookup = Xrm.Page.getAttribute("ownerid").getValue();
    var OwnerGuid = OwnerLookup[0].id;
    OwnerGuid = OwnerGuid.replace(/[{}]/g, "");
    var ContactLookUp = Xrm.Page.getAttribute("ccc_contact").getValue();
    var ContactId = ContactLookUp[0].id.replace(/[{}]/g, "");
    var contactTypeName = ContactLookUp[0].typename;
    var contactName = ContactLookUp[0].name;
    var signature = getSignature(OwnerGuid);
    if (Xrm.Page.getAttribute("ccc_recordurl") !== null && Xrm.Page.getAttribute("ccc_recordurl") !== undefined) {
        recordURL = Xrm.Page.getAttribute("ccc_recordurl").getValue();
        //alert(estimateNumber);
    }
    if (Xrm.Page.getAttribute("name") !== null && Xrm.Page.getAttribute("name") !== undefined) {
        estimateNumber = Xrm.Page.getAttribute("name").getValue();
        //alert(estimateNumber);
    }
    if (Xrm.Page.getAttribute("ccc_propertyid") !== null && Xrm.Page.getAttribute("ccc_propertyid") !== undefined) {
        var propertyNameRef = Xrm.Page.getAttribute("ccc_propertyid");
        if (propertyNameRef != null && propertyNameRef != undefined) {
            propertyId = propertyNameRef.getValue()[0].id.slice(1, -1);
            propertyAddress = propertyNameRef.getValue()[0].name;
            var object = getPropertyAddress(propertyId);
            propertyName = object[1];
            propertyCity = object[2];
        }
    }

    if (signature == null || signature == undefined) {
        signature = "";
    }
    var string = "Hello </br>Here is the estimate you requested for this location:</br>Estimate #  </br></br></br>A PDF copy is attached to this email, as well.</br>If you have any questions about this estimate, please reply back to this email or call me.Thank you for considering us!</br></br>" + signature;


    email["subject"] = "DonationRecipt";
    email["description"] = string;
    email["regardingobjectid_quote@odata.bind"] = "/aka_donation(" + qid + ")";
    //activityparty collection
    var activityparties = [];
    //from party
    var from = {};
    from["partyid_systemuser@odata.bind"] = "/systemusers(" + OwnerGuid + ")";
    from["participationtypemask"] = 1;
    //to party
    var to = {};
    to["partyid_contact@odata.bind"] = "/contacts(" + ContactId + ")";
    to["participationtypemask"] = 2;

    activityparties.push(to);
    activityparties.push(from);

    //set to and from to email
    email["email_activity_parties"] = activityparties;

    var req = new XMLHttpRequest();
    req.open("POST", Xrm.Page.context.getClientUrl() + "/api/data/v8.2/emails", false);
    req.setRequestHeader("OData-MaxVersion", "4.0");
    req.setRequestHeader("OData-Version", "4.0");
    req.setRequestHeader("Accept", "application/json");
    req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    req.onreadystatechange = function () {
        if (this.readyState === 4) {
            req.onreadystatechange = null;
            if (this.status === 204) {
                var uri = this.getResponseHeader("OData-EntityId");
                var regExp = /\(([^)]+)\)/;
                var matches = regExp.exec(uri);
                var newEntityId = matches[1];
                createEmailAttachment(newEntityId, base64);
            } else {
                Xrm.Utility.alertDialog(this.statusText);
            }
        }
    };
    req.send(JSON.stringify(email));
    ///////////////////
}
//function createEmailAttachment(emailUri, base64) {
 //   var activityId = emailUri.replace(/[{}]/g, "");
  //  var propertyId;
   // var propertyName;
  //  var propertyAddress;
  //  var propertyCity;
 //   var estimateNumber;
  //  if (Xrm.Page.getAttribute("name") !== null && Xrm.Page.getAttribute("name") !== undefined) {
   //     estimateNumber = Xrm.Page.getAttribute("name").getValue();
   // }
   // if (Xrm.Page.getAttribute("ccc_propertyid") !== null && Xrm.Page.getAttribute("ccc_propertyid") !== undefined) {
     //   var propertyNameRef = Xrm.Page.getAttribute("ccc_propertyid");
      //  if (propertyNameRef != null && propertyNameRef != undefined) {
          //  propertyId = propertyNameRef.getValue()[0].id.slice(1, -1);
          //  propertyAddress = propertyNameRef.getValue()[0].name;
          //  var object = getPropertyAddress(propertyId);
          //  propertyName = object[1];
          //  propertyCity = object[2];

      //  }
   // }


    var activityType = "email"; //or any other entity type
    var entity = {};
    entity["objectid_activitypointer@odata.bind"] = "/activitypointers(" + activityId + ")";
    //entity.body = "ZGZnZA=="; //your file encoded with Base64
    entity.body = base64; //your file encoded with Base64
    entity.filename = estimateNumber + "-" + propertyName + "-" + propertyAddress + "-" + propertyCity + ".pdf";
    entity.subject = estimateNumber + "-" + propertyName;
    entity.objecttypecode = activityType;
    var req = new XMLHttpRequest();
    req.open("POST", Xrm.Page.context.getClientUrl() + "/api/data/v8.2/activitymimeattachments", false);
    req.setRequestHeader("OData-MaxVersion", "4.0");
    req.setRequestHeader("OData-Version", "4.0");
    req.setRequestHeader("Accept", "application/json");
    req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    req.onreadystatechange = function () {
        if (this.readyState === 4) {
            req.onreadystatechange = null;
            if (this.status === 204) {
                var uri = this.getResponseHeader("OData-EntityId");
                var regExp = /\(([^)]+)\)/;
                var matches = regExp.exec(uri);
                var newEntityId = matches[1];
                //alert("attachement created "+newEntityId);
            } else {
                Xrm.Utility.alertDialog(this.statusText);
            }
        }
    };
    req.send(JSON.stringify(entity));
}
//function getPropertyAddress(pid) {
  //  debugger;
   // var ccc_name;
  //  var ccc_property1;
  //  var ccc_propertycity;
  //  var req = new XMLHttpRequest();
  //  req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v8.2/ccc_properties(" + pid + ")?//$select=ccc_name,ccc_property1,ccc_propertycity", false);
  //  req.setRequestHeader("OData-MaxVersion", "4.0");
  //  req.setRequestHeader("OData-Version", "4.0");
  //  req.setRequestHeader("Accept", "application/json");
  //  req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
   // req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
   // req.onreadystatechange = function () {
      //  if (this.readyState === 4) {
          //  req.onreadystatechange = null;
           // if (this.status === 200) {
             //   var result = JSON.parse(this.response);
              //  ccc_name = result["ccc_name"];
             //   ccc_property1 = result["ccc_property1"];
             //   ccc_propertycity = result["ccc_propertycity"];
           // } else {
              //  Xrm.Utility.alertDialog(this.statusText);
           // }
       // }
   // };
   // req.send();
   // return [ccc_name, ccc_property1, ccc_propertycity];
//}
function getSignature(OwnerGuid) {
    debugger;
    var sig;
    var req = new XMLHttpRequest();
    req.open("GET", Xrm.Page.context.getClientUrl() + "/api/data/v8.2/emailsignatures?$select=presentationxml&$filter=_ownerid_value eq " + OwnerGuid, false);
    req.setRequestHeader("OData-MaxVersion", "4.0");
    req.setRequestHeader("OData-Version", "4.0");
    req.setRequestHeader("Accept", "application/json");
    req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
    req.setRequestHeader("Prefer", "odata.include-annotations=\"*\"");
    req.onreadystatechange = function () {
        if (this.readyState === 4) {
            req.onreadystatechange = null;
            if (this.status === 200) {
                var results = JSON.parse(this.response);
                for (var i = 0; i < results.value.length; i++) {
                    var presentationxml = results.value[i]["presentationxml"];
                    oXml = CreateXmlDocument(presentationxml);
                    sig = oXml.lastChild.lastElementChild.textContent;
                }
            } else {
                Xrm.Utility.alertDialog(this.statusText);
            }
        }
    };
    req.send();
    return sig;
}
function CreateXmlDocument(signatureXmlStr) {
    // Function to create Xml formate of return email template data
    var parseXml;

    if (window.DOMParser) {
        parseXml = function (xmlStr) {
            return (new window.DOMParser()).parseFromString(xmlStr, "text/xml");
        };
    }
    else if (typeof window.ActiveXObject != "undefined" && new window.ActiveXObject("Microsoft.XMLDOM")) {
        parseXml = function (xmlStr) {
            var xmlDoc = new window.ActiveXObject("Microsoft.XMLDOM");
            xmlDoc.async = "false";
            xmlDoc.loadXML(xmlStr);

            return xmlDoc;
        };
    }
    else {
        parseXml = function () { return null; }
    }

    var xml = parseXml(signatureXmlStr);
    if (xml) {
        return xml;
    }
}