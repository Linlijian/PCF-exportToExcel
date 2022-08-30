import {IInputs, IOutputs} from "./generated/ManifestTypes";
import { Injectable } from '@angular/core';
//import * as XLSX from 'xlsx';
import * as XLSX from 'sheetjs-style'

const EXCEL_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';

export class excelexport implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	private button: HTMLButtonElement;
	private filename:string;
	private _notifyOutputChanged: () => void;
	private printable:string;
	private excel_style:string;
	private excel_header: string;
	private testss:JSON;

	/**
	 * Empty constructor.
	 */
	constructor()
	{

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		// Add control initialization code
		this.button = document.createElement("button");
		// Get the localized string from localized string
		if(context.parameters.pkButtonText.raw)
		{
			this.button.innerHTML = context.parameters.pkButtonText.raw;
		}
		this.button.classList.add("SimpleIncrement_Button_Style");
		this.button.classList.add("pkButton");
		//if(context.parameters.BackgroundColor.raw){
		//	this.button.style.backgroundColor = context.parameters.BackgroundColor.raw;
		//}; 
		if(context.parameters.pkTextColor.raw)
		{
			this.button.style.color = context.parameters.pkTextColor.raw;
		};
		if(context.parameters.pkTextSize.raw)
		{
			this.button.style.fontSize = context.parameters.pkTextSize.raw;
		};
		if(context.parameters.pkFont.raw)
		{
			this.button.style.fontFamily = context.parameters.pkFont.raw;
		};

		//new
		if (context.parameters.pkBorderRadius.raw) {
			this.button.style.borderRadius = context.parameters.pkBorderRadius.raw;;
		};
		if (context.parameters.pkPaddingRight.raw) {
			this.button.style.paddingRight = context.parameters.pkPaddingRight.raw;;
		};

		//solution v.2
		if (context.parameters.pkButtonColor.raw) {
			this.button.style.setProperty('--color', context.parameters.pkButtonColor.raw);
		};
		if (context.parameters.pkButtonColorHover.raw) {
			this.button.style.setProperty('--hover', context.parameters.pkButtonColorHover.raw);
		};

		this.button.style.textAlign = "right";

		this._notifyOutputChanged = notifyOutputChanged;
		//this.button.addEventListener("click", (event) => { this._value = this._value + 1; this._notifyOutputChanged();});
		this.button.addEventListener("click", this.onButtonClick.bind(this));
		// Adding the label and button created to the container DIV.
		this.button.style.width = container.style.width;
		if(context.parameters.pkButtonHeight.raw)
		{
			this.button.style.height = context.parameters.pkButtonHeight.raw + "px";
		};

		container.id = "containerid";
		container.appendChild(this.button);
	}


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		// Add code to update control view
		if(context.parameters.pkJsonformatData.raw)
		{
			this.printable = context.parameters.pkJsonformatData.raw;
		};
		if( context.parameters.pkFileName.raw)
		{
		this.filename = context.parameters.pkFileName.raw+".xlsx";
		};
		if(context.parameters.pkButtonText.raw)
		{
			this.button.innerHTML = context.parameters.pkButtonText.raw;
		};
		//if(context.parameters.BackgroundColor.raw){
		//	this.button.style.backgroundColor = context.parameters.BackgroundColor.raw;
		//}; 
		if(context.parameters.pkTextColor.raw)
		{
			this.button.style.color = context.parameters.pkTextColor.raw;
		};
		if(context.parameters.pkTextSize.raw)
		{
			this.button.style.fontSize = context.parameters.pkTextSize.raw;
		};
		if(context.parameters.pkFont.raw)
		{
			this.button.style.fontFamily = context.parameters.pkFont.raw;
		};
		if(context.parameters.pkButtonHeight.raw)
		{
			this.button.style.height = context.parameters.pkButtonHeight.raw + "px";
		};

		//new
		if (context.parameters.pkBorderRadius.raw) {
			this.button.style.borderRadius = context.parameters.pkBorderRadius.raw;;
		};
		if (context.parameters.pkPaddingRight.raw) {
			this.button.style.paddingRight = context.parameters.pkPaddingRight.raw;;
		};

		//solution v.2
		if (context.parameters.pkButtonColor.raw) {
			this.button.style.setProperty('--color', context.parameters.pkButtonColor.raw);
		};
		if (context.parameters.pkButtonColorHover.raw) {
			this.button.style.setProperty('--hover', context.parameters.pkButtonColorHover.raw);
		};

		//solution v2.1
		this.excel_header = context.parameters.pkExcelHeader.raw!;
		this.excel_style = context.parameters.pkExcelStyle.raw!;
	}

	private onButtonClick(event: Event): void {
		if (this.excel_header && this.excel_style) {
			let row2 = JSON.parse(this.excel_style)
			let Heading = [JSON.parse(this.excel_header)];

			const wb = XLSX.utils.book_new();
			const ws: XLSX.WorkSheet = XLSX.utils.json_to_sheet([]);

			XLSX.utils.sheet_add_aoa(ws, Heading);
			XLSX.utils.sheet_add_json(ws, row2, { origin: 'A2', skipHeader: true });
			XLSX.utils.book_append_sheet(wb, ws, 'FailSheet');
			XLSX.writeFile(wb, this.filename);
		} else {
			var testss = JSON.parse(this.printable);
			const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(testss);
			const workbook: XLSX.WorkBook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
			XLSX.writeFile(workbook, this.filename);
        }
	}

	public ex_color() {
		// STEP 1: Create a new workbook
		const wb = XLSX.utils.book_new();

		// STEP 2: Create data rows and styles
		let row = [
			{ v: "Courier: 24", t: "s", s: { font: { name: "Courier", sz: 24 } } },
			{ v: "bold & color", t: "s", s: { font: { bold: true, color: { rgb: "FF0000" } } } },
			{ v: "fill: color", t: "s", s: { fill: { fgColor: { rgb: "E9E9E9" } } } },
			{ v: "line\nbreak", t: "s", s: { alignment: { wrapText: true } } }
		];

		// STEP 3: Create worksheet with rows; Add worksheet to workbook
		const ws = XLSX.utils.aoa_to_sheet([row]);
		XLSX.utils.book_append_sheet(wb, ws, "readme demo");

		// STEP 4: Write Excel file to browser
		XLSX.writeFile(wb, "xlsx-js-style-demo.xlsx");


		//---
		var workbook2 = XLSX.utils.book_new();

		var ws2 = XLSX.utils.aoa_to_sheet([
			["A1", "B1", "C1"],
			["A2", "B2", "C2"],
			["A3", "B3", "C3"]
		])
		ws2['A1'].s = {
			font: {
				name: 'arial',
				sz: 24,
				bold: true,
				color: "#F2F2F2"
			},
			fill: {
				fgColor: { rgb: "E9E9E9" }
			}
		}
		ws2['A3'].s = {
			font: {
				name: 'arial',
				sz: 24,
				bold: true,
				color: "#F2F2F2"
			},
		}

		XLSX.utils.book_append_sheet(workbook2, ws2, "SheetName");
		XLSX.writeFile(workbook2, 'FileName.xlsx');

    }
	public xx() {
		let arr = [
			{ firstName: 'Jack', lastName: 'Sparrow', email: 'abc@example.com' },
			{ firstName: 'Harry', lastName: 'Potter', email: 'abc@example.com', aa:"saas" },
		];

		let Heading = [['FirstName', 'Last Name', 'Email']];

		//Had to create a new workbook and then add the header
		const wb = XLSX.utils.book_new();
		const ws: XLSX.WorkSheet = XLSX.utils.json_to_sheet([]);
		XLSX.utils.sheet_add_aoa(ws, Heading);

		//Starting in the second row to avoid overriding and skipping headers
		XLSX.utils.sheet_add_json(ws, arr, { origin: 'A2', skipHeader: true });

		XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

		XLSX.writeFile(wb, 'filename2.xlsx');
	}
	public export_but_not_have_header() {
		// STEP 1: Create a new workbook
		const wb = XLSX.utils.book_new();

		// STEP 2: Create data rows and styles
		let row = [
			{ v: "Courier: 24", t: "s", s: { font: { name: "Courier", sz: 24 } } },
			{ v: "bold & color", t: "s", s: { font: { bold: true, color: { rgb: "FF0000" } } } },
			{ v: "fill: color", t: "s", s: { fill: { fgColor: { rgb: "E9E9E9" } } } },
			{ v: "line\nbreak", t: "s", s: { alignment: { wrapText: true } } }
		];

		let row2 = [
			[
				{ v: "Courier: 24", t: "s", s: { font: { name: "Courier", sz: 24 } } },
				{ v: "bold & color", t: "s", s: { font: { bold: true, color: { rgb: "FF0000" } } } },
			],
			[
				{ v: "2: 24", t: "s", s: { font: { name: "Courier", sz: 24 } } },
				{ v: "4444 & color", t: "s", s: { font: { bold: true, color: { rgb: "FF0000" } } } },
			],
		]

		let Heading = [['FirstName', 'Last Name']];

		// STEP 3: Create worksheet with rows; Add worksheet to workbook
		const ws = XLSX.utils.aoa_to_sheet(row2);
		XLSX.utils.book_append_sheet(wb, ws, "readme demo");

		// STEP 4: Write Excel file to browser
		XLSX.writeFile(wb, "xlsx-js-style-demo.xlsx");
		console.log(row2)
		console.log('---------------------')
		console.log(ws)
		console.log('---------------------')
		console.log(wb)
		console.log('---------------------')
	}
	public export_have_header() {
		// STEP 1: Create a new workbook
		const wb = XLSX.utils.book_new();

		// STEP 2: Create data rows and styles

		let row2 = [
			[
				{ v: "Courier: 24", t: "s", s: { font: { name: "Courier", sz: 24 } } },
				{ v: "bold & color", t: "s", s: { font: { bold: true, color: { rgb: "FF0000" } } } },
			],
			[
				{ v: "2: 24", t: "s", s: { font: { name: "Courier", sz: 24 } } },
				{ v: "4444 & color", t: "s", s: { fill: { fgColor: { rgb: "E9E9E9" } } } },
			],
		]

		let Heading = [['FirstName', 'Last Name']];

		// STEP 3: Create worksheet with rows; Add worksheet to workbook
		//const ws = XLSX.utils.aoa_to_sheet(row2);
		//XLSX.utils.book_append_sheet(wb, ws, "readme demo");


		//Had to create a new workbook and then add the header
		//const wb = XLSX.utils.book_new();
		const ws: XLSX.WorkSheet = XLSX.utils.json_to_sheet([]);
		XLSX.utils.sheet_add_aoa(ws, Heading);

		//Starting in the second row to avoid overriding and skipping headers
		XLSX.utils.sheet_add_json(ws, row2, { origin: 'A2', skipHeader: true });

		XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

		//XLSX.writeFile(wb, 'filename2.xlsx');

		// STEP 4: Write Excel file to browser
		XLSX.writeFile(wb, "header.xlsx");
		console.log(row2)
		console.log('---------------------')
		console.log(ws)
		console.log('---------------------')
		console.log(wb)
		console.log('---------------------')
	}
	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		return {};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		// Add code to cleanup control if necessary
	}
}