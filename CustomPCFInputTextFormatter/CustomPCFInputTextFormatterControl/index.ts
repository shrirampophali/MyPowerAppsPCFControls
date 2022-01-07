import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class CustomPCFInputTextFormatterControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	private _notifyOutputChanged: () => void;
	private _formatTypePropertyValue: string;
	private _customFormatTextPropertyValue: string;
	private _customFormatTextMessagePropertyValue: string;
	private _mandatoryPropertyValue: boolean;
	private _inputBoxCssStylePropertyValue: string;
	private _inputBoxRequiredCssStylePropertyValue: string;
	private _outputValue: string;
	private _isValidValue: boolean;
	private _inputElement: HTMLInputElement;
	private _container: HTMLDivElement;
	private _rootcontainer: HTMLDivElement;
	private _refreshData: EventListenerOrEventListenerObject;

	/*constructor()
	{

	}*/

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {

		// PCF Input Formatter Control
		this._notifyOutputChanged = notifyOutputChanged;
		this._refreshData = this.refreshData.bind(this);
		this._rootcontainer = document.createElement("div");
		this._container = document.createElement("div");
		this._container.setAttribute("id", "divinputfield");

		this._inputElement = document.createElement("input");
		this._inputElement.setAttribute("id", "pcfinputformatterid");
		this._inputElement.addEventListener("input", this._refreshData);

		// Store values so as to return to parent app via output variables (if needed)
		this._formatTypePropertyValue = context.parameters.formatTypeProperty.raw!;
		this._customFormatTextPropertyValue = context.parameters.CustomFormatTextProperty.raw!;
		this._customFormatTextMessagePropertyValue = context.parameters.CustomFormatTextMessageProperty.raw!;
		this._mandatoryPropertyValue = context.parameters.MandatoryProperty.raw!;
		this._inputBoxCssStylePropertyValue = context.parameters.InputBoxCssStyleProperty.raw!;
		this._inputBoxRequiredCssStylePropertyValue = context.parameters.InputBoxRequiredCssStyleProperty.raw!;

		this._inputElement.setAttribute("style", this._inputBoxCssStylePropertyValue);

		// Render control based on Format Type		
		if (this._formatTypePropertyValue == "Date") {
			this._inputElement.setAttribute("type", "date");
		}
		if (this._formatTypePropertyValue == "Time") {
			this._inputElement.setAttribute("type", "time");
		}
		if (this._formatTypePropertyValue == "DateTime") {
			this._inputElement.setAttribute("type", "datetime-local");
		}
		if (this._formatTypePropertyValue == "Password") {
			this._inputElement.setAttribute("type", "password");
			this._inputElement.setAttribute("title", this._customFormatTextMessagePropertyValue);
			this._inputElement.setAttribute("placeholder", this._customFormatTextMessagePropertyValue);
		}
		if (this._formatTypePropertyValue == "Custom") {
			// if custom format type
			this._inputElement.setAttribute("type", "text");
			this._inputElement.setAttribute("pattern", this._customFormatTextPropertyValue);
			this._inputElement.setAttribute("title", this._customFormatTextMessagePropertyValue);
			this._inputElement.setAttribute("placeholder", this._customFormatTextMessagePropertyValue);
		}

		this._inputElement.setAttribute("style", this._inputBoxCssStylePropertyValue);

		// if field is mandatory
		if (this._mandatoryPropertyValue == true) {

			this._inputElement.setAttribute("required", "required");
			this._inputElement.setAttribute("style", this._inputBoxRequiredCssStylePropertyValue);
		}

		// To show stored value to model-driven or canvas app form field (if associated with any form field)
		try { this._inputElement.value = context.parameters.OutputValue.formatted ? context.parameters.OutputValue.formatted : ""; }
		catch { }

		this._container.appendChild(this._inputElement);
		this._rootcontainer.appendChild(this._container);
		container.appendChild(this._rootcontainer);
	}

	// Triggers when any property value is updated
	public updateView(context: ComponentFramework.Context<IInputs>): void {

		// Store values so as to return to parent app via output variables (if needed)
		this._formatTypePropertyValue = context.parameters.formatTypeProperty.raw!;
		this._customFormatTextPropertyValue = context.parameters.CustomFormatTextProperty.raw!;
		this._customFormatTextMessagePropertyValue = context.parameters.CustomFormatTextMessageProperty.raw!;
		this._mandatoryPropertyValue = context.parameters.MandatoryProperty.raw!;
		this._inputBoxCssStylePropertyValue = context.parameters.InputBoxCssStyleProperty.raw!;
		this._inputBoxRequiredCssStylePropertyValue = context.parameters.InputBoxRequiredCssStyleProperty.raw!;

		// Render control based on Format Type		
		if (this._formatTypePropertyValue == "Date") {
			this._inputElement.setAttribute("type", "date");
		}
		if (this._formatTypePropertyValue == "DateTime") {
			this._inputElement.setAttribute("type", "datetime-local");
		}
		if (this._formatTypePropertyValue == "Time") {
			this._inputElement.setAttribute("type", "time");
		}
		if (this._formatTypePropertyValue == "Password") {
			this._inputElement.setAttribute("type", "password");
			this._inputElement.setAttribute("title", this._customFormatTextMessagePropertyValue);
			this._inputElement.setAttribute("placeholder", this._customFormatTextMessagePropertyValue);
		}
		if (this._formatTypePropertyValue == "Custom") {
			// if custom format type
			this._inputElement.setAttribute("type", "text");
			this._inputElement.setAttribute("pattern", this._customFormatTextPropertyValue);
			this._inputElement.setAttribute("title", this._customFormatTextMessagePropertyValue);
			this._inputElement.setAttribute("placeholder", this._customFormatTextMessagePropertyValue);
		}

		// Validation start
		this._isValidValue = false;
		this._outputValue = (this._inputElement.value as any) as string;
		this._inputElement.setAttribute("style", this._inputBoxCssStylePropertyValue);
		if (this._mandatoryPropertyValue == true) {

			if (this._formatTypePropertyValue == "Custom") {

				// If value is invalid
				this._inputElement.setAttribute("style", this._inputBoxRequiredCssStylePropertyValue);

				// if custom format type, validate with regular expression
				if (this._customFormatTextPropertyValue.trim() != "") {

					this._inputElement.setAttribute("type", "text");
					this._inputElement.setAttribute("pattern", this._customFormatTextPropertyValue);
					this._inputElement.setAttribute("title", this._customFormatTextMessagePropertyValue);
					this._inputElement.setAttribute("placeholder", this._customFormatTextMessagePropertyValue);

					var regexp = new RegExp(this._customFormatTextPropertyValue);
					var test = regexp.test(this._outputValue);

					if (!test) {
						// If value is invalid
						this._isValidValue = false;
						this._inputElement.setAttribute("style", this._inputBoxRequiredCssStylePropertyValue);
					}
					else {

						// If value is valid
						this._isValidValue = true;
						this._inputElement.setAttribute("style", this._inputBoxCssStylePropertyValue);
					}
				}
			}
			else { // Other Format Type than Custom
				if (this._outputValue.trim() == "") {
					// If value is invalid
					this._isValidValue = false;
					this._inputElement.setAttribute("style", this._inputBoxRequiredCssStylePropertyValue);
				}
				else {
					// If value is valid
					this._isValidValue = true;
					this._inputElement.setAttribute("style", this._inputBoxCssStylePropertyValue);
				}
			}
		}
		else {
			// if not mandatory, Custom Pattern and has some input value
			if (this._formatTypePropertyValue == "Custom" && this._outputValue.trim() != "") {

				// if custom format type, validate with regular expression
				if (this._customFormatTextPropertyValue.trim() != "") {

					this._inputElement.setAttribute("type", "text");
					this._inputElement.setAttribute("pattern", this._customFormatTextPropertyValue);
					this._inputElement.setAttribute("title", this._customFormatTextMessagePropertyValue);
					this._inputElement.setAttribute("placeholder", this._customFormatTextMessagePropertyValue);

					var regexp = new RegExp(this._customFormatTextPropertyValue);
					var test = regexp.test(this._outputValue);

					if (!test) {
						// If value is invalid
						this._isValidValue = false;
						this._inputElement.setAttribute("style", this._inputBoxRequiredCssStylePropertyValue);
					}
					else {
						// If value is valid
						this._isValidValue = true;
						this._inputElement.setAttribute("style", this._inputBoxCssStylePropertyValue);
					}
				}
			}
		}
		// Validation end
	}

	public getOutputs(): IOutputs {
		return {
			OutputValue: this._outputValue,
			IsValid: this._isValidValue
		};
	}

	// Triggers when any input is provided in input text box
	public refreshData(evt: Event): void {

		// Validation start
		this._isValidValue = false;
		this._outputValue = (this._inputElement.value as any) as string;
		this._inputElement.setAttribute("style", this._inputBoxCssStylePropertyValue);
		if (this._mandatoryPropertyValue == true) {

			if (this._formatTypePropertyValue == "Custom") {

				// If value is invalid
				this._inputElement.setAttribute("style", this._inputBoxRequiredCssStylePropertyValue);

				// if custom format type, validate with regular expression
				if (this._customFormatTextPropertyValue.trim() != "") {

					this._inputElement.setAttribute("type", "text");
					this._inputElement.setAttribute("pattern", this._customFormatTextPropertyValue);
					this._inputElement.setAttribute("title", this._customFormatTextMessagePropertyValue);
					this._inputElement.setAttribute("placeholder", this._customFormatTextMessagePropertyValue);

					var regexp = new RegExp(this._customFormatTextPropertyValue);
					var test = regexp.test(this._outputValue);

					if (!test) {
						// If value is invalid
						this._isValidValue = false;
						this._inputElement.setAttribute("style", this._inputBoxRequiredCssStylePropertyValue);
					}
					else {
						// If value is valid
						this._isValidValue = true;
						this._inputElement.setAttribute("style", this._inputBoxCssStylePropertyValue);
					}
				}
			}
			else { // Other Format Type than Custom
				if (this._outputValue.trim() == "") {
					// If value is invalid
					this._isValidValue = false;
					this._inputElement.setAttribute("style", this._inputBoxRequiredCssStylePropertyValue);
				}
				else {
					// If value is valid
					this._isValidValue = true;
					this._inputElement.setAttribute("style", this._inputBoxCssStylePropertyValue);
				}
			}
		}
		else {
			// if not mandatory, Custom Pattern and has some input value
			if (this._formatTypePropertyValue == "Custom" && this._outputValue.trim() != "") {

				// if custom format type, validate with regular expression
				if (this._customFormatTextPropertyValue.trim() != "") {

					this._inputElement.setAttribute("type", "text");
					this._inputElement.setAttribute("pattern", this._customFormatTextPropertyValue);
					this._inputElement.setAttribute("title", this._customFormatTextMessagePropertyValue);
					this._inputElement.setAttribute("placeholder", this._customFormatTextMessagePropertyValue);

					var regexp = new RegExp(this._customFormatTextPropertyValue);
					var test = regexp.test(this._outputValue);

					if (!test) {
						// If value is invalid
						this._isValidValue = false;
						this._inputElement.setAttribute("style", this._inputBoxRequiredCssStylePropertyValue);
					}
					else {
						// If value is valid
						this._isValidValue = true;
						this._inputElement.setAttribute("style", this._inputBoxCssStylePropertyValue);
					}
				}
			}
		}
		// Validation end

		// Notify once output is changed
		this._notifyOutputChanged();
	}

	public destroy(): void {
	}
}
