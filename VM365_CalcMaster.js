/* global _spPageContextInfo */
/* global document,console,$ */
var VM365_CalcMaster = {
	//global settings
	comma: _spPageContextInfo.currentLanguage === 1033 ? ',' : ';',
	layouts: _spPageContextInfo.layoutsUrl,
	textareaWidth: '600px',
	textareaHeight: '300px',
	//initialize CalcMaster
	initialize: function () {
		VM365_CalcMaster.columnList = document.getElementById('onetidIOCalcFields1'); //get references to DOM elements
		VM365_CalcMaster.textarea = document.getElementById('onetidIODefTextValue1');
		//adjust layout of textarea and column list
		VM365_CalcMaster.columnList.style.height =
		VM365_CalcMaster.textarea.style.height = VM365_CalcMaster.textareaHeight;
		VM365_CalcMaster.textarea.style.width = VM365_CalcMaster.textareaWidth;
		//Highlight the Number datatype for HTML output
		document.getElementById('L_onetidTypeNumber1').innerHTML = 'Number (outputs HTML/JavaScript)';
		document.getElementById('L_onetidTypeNumber1').parentNode.style.backgroundColor = 'lightGreen';
		//add CalcMaster UI elements for Toolbar and Reporting
		document.getElementById('onetidIODefText0').innerHTML = "<div style='min-height:45px;'><div id='calcmasterSave'></div><div id='calcmasterHint'></div><div id='calcmasterButtons'></div></div>";
		//create CalcMaster toolbar HTML
		var H = "<img src='/" + VM365_CalcMaster.layouts + "/images/WSC16.GIF' ";
			H += " alt='Wrap Javascript code in SharePoint Formula quoted string' ";
			H += " onclick='VM365_CalcMaster.functionConvertToICC()'>Convert JS";
		VM365_CalcMaster.setDOMelement(H, 0, 'calcmasterButtons');
		//add event to the existing textarea
		VM365_CalcMaster.textarea.onkeyup = function (event) {
			if ([37, 38, 39, 40].indexOf(event.keyCode) === -1) { //ignore arrow keys
				VM365_CalcMaster.updateFormula(event);
			}
		};
	},
	//supporting functions
	setDOMelement: function (txt, error, id) {
		var savestate = document.getElementById( id || 'calcmasterSave' );
		savestate.innerHTML = txt;
		savestate.style.color = ['black', 'red', 'green', 'darkorange'][error];
	},
	setFormulaInTextarea: function (Formula) {
		if (Formula.indexOf('\n') === -1) {
			Formula = Formula.replace(/&\"/gi, '\n  &\"'); //replace & before quote
			Formula = Formula.replace(/\"&/gi, '"\n  &'); //replace & after quote
			Formula = Formula.replace(/\";/gi, '\"\n;'); //replace & before quote
		}
		VM365_CalcMaster.textarea.value = Formula;
	},
	sanitizeFormulaBeforeWritingToSP: function (Formula) {
		//F=F.replace(/'/g, "&#39;");
		//Formula = Formula.replace(/'/g, "\\'");
		//F=F.replace(/'/g, '"&CHAR(39)&"');
		Formula = Formula.replace(/`/g, "'"); //replace backticks with single quote
		return (Formula);
	},
	//addtional button functions while editting
	functionConvertToICC: function (Formula) {
		Formula = Formula || VM365_CalcMaster.textarea.value;
		if (Formula[0] !== '=') {//if Formula is NOT a SharePoint formula
			//it is javascript code
			Formula.replace(/'/g, '`');
			if (Formula.indexOf('onload') === -1 && Formula.indexOf('onclick') === -1) {
				//wrap in IMG onload tag
				Formula = '="<img src=/_layouts/images/blank.gif onload=""{"&"' + Formula.replace(/\n/g, '"\n&"') + '"&"}"">"';
			}
			VM365_CalcMaster.setFormulaInTextarea(Formula);
		}
	},
	updateFormula: function (event) {
		//current selected Formula output datatype
		var selectedType = Array.prototype.filter.call(document.getElementsByName('ResultType'), function (r) {
			if (r.checked) return (r);
		});
		document.getElementById('onetidSaveItem').style.display = 'none'; //hide OK button 
		var Formula = VM365_CalcMaster.textarea.value;
		VM365_CalcMaster.setFormulaInTextarea(Formula);
		VM365_CalcMaster.updateCalculatedColumn({
			Title: document.getElementById('idColName').value,
			Description: document.getElementById('idDesc').value,
			Formula: Formula,
			type: selectedType[0].value //Text or Number or DateTime or Currency or Boolean
		});
		//reporting of Formula analysis
		var calcmasterHint = [];
		if (Formula.split('"').length % 2 === 0) calcmasterHint.push("unmatched double quotes");
		if (Formula.split("'").length % 2 === 0) calcmasterHint.push("unmatched single quotes");
		if (Formula.split('(').length !== Formula.split(')').length) calcmasterHint.push("unmatched () brackets");
		VM365_CalcMaster.setDOMelement(calcmasterHint.join('<br>'), 3, 'calcmasterHint');
	},
	updateCalculatedColumn: function (column) {
		VM365_CalcMaster.setDOMelement('Formula: updating...', 2);
		column.list = "Lists(guid'"   + _spPageContextInfo.pageListId.replace(/[{}]/g, '') +   "')";
		column.field = "Fields/getbytitle('" + column.Title + "')";
		var REST = {
			url: _spPageContextInfo.webAbsoluteUrl + "/_api/Web/" + column.list + "/" + column.field,
			success: function (data) {
				VM365_CalcMaster.setDOMelement('Formula: saved', 0);
			},
			error: function (data) {
				VM365_CalcMaster.setDOMelement(data.responseJSON.error.message.value, 1);
			},
			type: "POST",
			headers: {
				"X-HTTP-Method": "MERGE",
				"X-RequestDigest": document.getElementById('__REQUESTDIGEST').value,
				"accept": "application/json;odata=verbose",
				"content-type": "application/json;odata=verbose"
			},
			data: JSON.stringify({
				'Title': column.Title,
				'Formula': VM365_CalcMaster.sanitizeFormulaBeforeWritingToSP(column.Formula),
				'__metadata': {
					'type': 'SP.FieldCalculated'
				},
				'Description': column.Description,
				'OutputType': ({
					Text: 2,
					Number: 9,
					Currency: 10,
					DateTime: 4,
					Boolean: 8
				})[column.type]
			})
		};
		$.ajax(REST);
	}
};
VM365_CalcMaster.initialize();
