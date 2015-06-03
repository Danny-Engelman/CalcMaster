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
		var CM=VM365_CalcMaster;//pointer to global
		CM.columnList = document.getElementById('onetidIOCalcFields1'); //get references to DOM elements
		CM.textarea = document.getElementById('onetidIODefTextValue1');
		//adjust layout of textarea and column list
		CM.columnList.style.height =
		CM.textarea.style.height = CM.textareaHeight;
		CM.textarea.style.width = CM.textareaWidth;
		//Highlight the Number datatype for HTML output
		document.getElementById('L_onetidTypeNumber1').innerHTML = 'Number (outputs HTML/JavaScript)';
		document.getElementById('L_onetidTypeNumber1').parentNode.style.backgroundColor = 'lightGreen';
		//add CalcMaster UI elements for Toolbar and Reporting
		document.getElementById('onetidIODefText0').innerHTML = "<div style='min-height:45px;'><div id='calcmasterSave'></div><div id='calcmasterHint'></div><div id='calcmasterButtons'></div></div>";
		//create CalcMaster toolbar HTML
		var H = "<img src='/" + CM.layouts + '/images/WSC16.GIF';
			H += " alt='Wrap Javascript code in SharePoint Formula quoted string' ";
			H += "' onclick='CM.functionConvertToICC()'>";
		CM.setDOMelement(H, 0, 'calcmasterButtons');
		//add event to the existing textarea
		CM.textarea.onkeyup = function (event) {
			if ([37, 38, 39, 40].indexOf(event.keyCode) === -1) { //ignore arrow keys
				CM.updateFormula(event);
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
		var CM=VM365_CalcMaster;//pointer to global
		if (Formula.indexOf('\n') === -1) {
			Formula = Formula.replace(/&\"/gi, '\n  &\"'); //replace & before quote
			Formula = Formula.replace(/\"&/gi, '"\n  &'); //replace & after quote
			Formula = Formula.replace(/\";/gi, '\"\n;'); //replace & before quote
		}
		CM.textarea.value = Formula;
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
		var CM=VM365_CalcMaster;//pointer to global
		Formula = Formula || CM.textarea.value;
		if (Formula[0] !== '=') {//if Formula is NOT a SharePoint formula
			//it is javascript code
			Formula.replace(/'/g, '`');
			if (Formula.indexOf('onload') === -1 && Formula.indexOf('onclick') === -1) {
				//wrap in IMG onload tag
				Formula = '="<img src=/_layouts/images/blank.gif onload=""{"&"' + Formula.replace(/\n/g, '"\n&"') + '"&"}"">"';
			}
			CM.setFormulaInTextarea(Formula);
		}
	},
	updateFormula: function (event) {
		var CM=VM365_CalcMaster;//pointer to global
		//current selected Formula output datatype
		var selectedType = Array.prototype.filter.call(document.getElementsByName('ResultType'), function (r) {
			if (r.checked) return (r);
		});
		document.getElementById('onetidSaveItem').style.display = 'none'; //hide OK button 
		var Formula = CM.textarea.value;
		CM.setFormulaInTextarea(Formula);
		CM.updateCalculatedColumn({
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
		CM.setDOMelement(calcmasterHint.join('<br>'), 3, 'calcmasterHint');
	},
	updateCalculatedColumn: function (column) {
		var CM=VM365_CalcMaster;//pointer to global
		CM.setDOMelement('Formula: updating...', 2);
		column.list = "Lists(guid'"   + _spPageContextInfo.pageListId.replace(/[{}]/g, '') +   "')";
		column.field = "Fields/getbytitle('" + column.Title + "')";
		var REST = {
			url: _spPageContextInfo.webAbsoluteUrl + "/_api/Web/" + column.list + "/" + column.field,
			success: function (data) {
				CM.setDOMelement('Formula: saved', 0);
			},
			error: function (data) {
				CM.setDOMelement(data.responseJSON.error.message.value, 1);
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
				'Formula': CM.sanitizeFormulaBeforeWritingToSP(column.Formula),
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
