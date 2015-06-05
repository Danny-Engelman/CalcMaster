/*global _spPageContextInfo */
/*global document,console,clearTimeout,setTimeout,$ */
/*jshint -W043*/
var VM365_CalcMaster = { //global obect for easier debugging
	//global settings
	comma: _spPageContextInfo.currentLanguage === 1033 ? ',' : ';', //SharePoint function separator
	layouts: _spPageContextInfo.layoutsUrl,
	saveID:0,
	//initialize CalcMaster
	initialize: function () {
		//get references to DOM elements
		VM365_CalcMaster.columnList = document.getElementById('onetidIOCalcFields1');
		VM365_CalcMaster.textarea = document.getElementById('onetidIODefTextValue1');
		//adjust layout of textarea and column list
		VM365_CalcMaster.columnList.style.height = VM365_CalcMaster.textarea.style.height = '300px';
		VM365_CalcMaster.textarea.style.width = '600px';
		//Highlight the Number datatype for HTML output
		VM365_CalcMaster.setDOMelement('Number (outputs HTML/JavaScript)', 0, 'L_onetidTypeNumber1', 'lightGreen');
		//add CalcMaster UI elements for Toolbar and Save Formula Reporting, remember the MIT license
		VM365_CalcMaster.setDOMelement(
			"<h2>CalcMaster</h2> \
			powered by <a target=_new href=http://365coach.nl>365Coach</a> - \
			SharePoint productivity enhancements \
			<br>See what you can do with OOB Calculated Columns at \
				<a href=http://viewmaster365.com/#/How>ViewMaster365.com</a> \
			<br>Also check out the \
				<a href='http://viewmaster365.com/365coach/#/Drag_Drop_Columns_in_the_EditView_Page'> \
					Drag-Drop View Editor</a> \
			<div style='min-height:55px;max-width:600px;'>\
				<img src='/" + VM365_CalcMaster.layouts + "/images/WSC16.GIF' \
					title='Wrap Javascript code in SharePoint Formula quoted string' \
					onclick='VM365_CalcMaster.functionConvertToICC()'> \
				<div id='calcmasterSave'></div><div id='calcmasterHint'></div> \
			</div>", 0, 'onetidIODefText0');
		VM365_CalcMaster.setDOMelement('Update your Formula and I will check if the Formula is correct',2);
		//add event to the existing textarea
		VM365_CalcMaster.textarea.onkeyup = function (event) {
			if ([37, 38, 39, 40].indexOf(event.keyCode) === -1) { //ignore arrow keys
				clearTimeout( VM365_CalcMaster.saveID );
				VM365_CalcMaster.saveID=setTimeout(function(){
								VM365_CalcMaster.updateFormula(event);
					},500);
			}
		};
	},
	//supporting functions
	setDOMelement: function (txt, color, id, bgcolor) {
		var element = document.getElementById(id || 'calcmasterSave');
		element.innerHTML = txt;
		element.style.color = ['black', 'red', 'green', 'darkorange'][color];
		if (bgcolor) element.style.backgroundColor = bgcolor;
	},
	setFormulaInTextarea: function (Formula) {
		if (Formula.indexOf('\n') === -1) { //if no newlines exist reformat to multiple lines
			Formula = Formula.replace(/&\"/gi, '\n  &\"'); //replace & before quote
			Formula = Formula.replace(/\"&/gi, '"\n  &'); //replace & after quote
			Formula = Formula.replace(/\";/gi, '\"\n;'); //replace & before quote
		}
		VM365_CalcMaster.textarea.value = Formula;
	},
	sanitizeFormulaBeforeWritingToSP: function (Formula) {
		//the original Grand CalcMaster does a lot more, left the code in as a reminder to self
		//Formula = Formula.replace(/'/g, "&#39;");
		//Formula = Formula.replace(/'/g, "\\'");
		//Formula = Formula.replace(/'/g, '"&CHAR(39)&"');
		Formula = Formula.replace(/`/g, "'"); //replace backticks with single quote
		return (Formula);
	},
	//additional button functions in CalcMaster editor
	functionConvertToICC: function (Formula) {
		Formula = Formula || VM365_CalcMaster.textarea.value;
		if (Formula[0] !== '=') { //if Formula is NOT a SharePoint formula, it is javascript code
			Formula.replace(/'/g, '`'); //replace original single quotes with backticks (used in Grand CalcMaster)
			if (Formula.indexOf('onload') === -1 && Formula.indexOf('onclick') === -1) {
				//wrap in IMG onload tag and quoted SharePoint formula strings
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
		VM365_CalcMaster.setDOMelement('Updating Formula ' + column.Title + '...', 2);
		column.list = "Lists(guid'" + _spPageContextInfo.pageListId.replace(/[{}]/g, '') + "')";
		column.field = "Fields/getbytitle('" + column.Title + "')";
		$.ajax({
			url: _spPageContextInfo.webAbsoluteUrl + "/_api/Web/ " + column.list + "/" + column.field,
			success: function (data) {
				VM365_CalcMaster.setDOMelement('Formula ' + column.Title + ' saved succesfully');
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
		}); //ajax call
	}
};
VM365_CalcMaster.initialize();
