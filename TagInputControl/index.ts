import * as Tagify from '@yaireo/tagify';
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { TagManagerService } from './TagManagerService';
import { Tag } from './Tag';
import { CustomError } from './ErrorDefinitions';
import { ApplicationInsights, IExceptionTelemetry } from '@microsoft/applicationinsights-web';
import './ExtensionMethods';

/**
 * Tag Input Control Class, implementation of Component Framework Standard Control.
 * There is a DI with the class @see TagManagerService
 */
export class TagInputControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	// the main input HTML control used by the Tagify component
	private _input: HTMLInputElement;

	// the PCF component context
	private _context: ComponentFramework.Context<IInputs>;

	// the main HTML element which has the input
	private _container: HTMLDivElement;

	// the HTML Tag element (<tag />)
	private _tagEl: HTMLElement;

	// the instance of the Tagify control
	private _tagify: Tagify;

	// readonly flag
	private _isDisabled: Boolean;

	// visible flag
	private _isVisible: Boolean;

	// the main service of this Component.
	private _tagManagerService: TagManagerService;

	// defines if the data was loaded. This will prevent the onAdd to be trigger during the first load
	private _initialized: boolean = false;

	// flag to be used to prevent any Tag to be added in case an error ocurred 
	private _currentState: string = '';

	// holds the instance to app insights (Azure)
	private _appInsights: ApplicationInsights;

	// accessor to the Resources file / strings
	private _resx: ComponentFramework.Resources;

	// defines the threshold to start querying the CDS data.
	private _searchActivationTreshold: Number;

	/**
	 * Empty constructor.
	 */
	constructor() {

	}

	/**
	 * @todo: Add a required validation rule and prevent user from saving the Model-Driven App form in case the user didn't provide the minimum of tags.
	 * @todo: Cover update method in order to allow an user to change the Tag name and its Translation's name.
	 * 
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Don't forget to @see TagManagerService class!
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {

		this._context = context;
		this._container = container;
		this._resx = context.resources;
		this._input = document.createElement("input");
		this._container.appendChild(this._input);

		let contextData: any = (<any>context.mode).contextInfo || {};

		this._tagManagerService = new TagManagerService(context.webAPI,
			context.parameters.Scope.raw || "default",
			context.parameters.DefaultLanguageLCID.raw || (window.navigator.language || 'en-us'), // window.navigator.language
			contextData.entityTypeName,
			contextData.entityId,
			`${window.location.protocol}//${window.location.host}`,
			context.parameters.ManyToManyRelationshipName.raw || '',
			context.parameters.IntersectEntityName.raw || '',
			context.parameters.WebApiBasePath.raw || '/api/data/v9.1',
			context.resources);

		this._searchActivationTreshold = context.parameters.AutocompleteActivationThreshold.raw || 2;

		this._tagify = new Tagify(this._input, {
			pattern: context.parameters.TagInputRegexPattern.raw || /^.{2,50}$/, 
			maxTags: context.parameters.AutocompleteMaxItems.raw || 20,
			enforceWhitelist: context.parameters.EnforceWhitelist.raw || false,
			addTagOnBlur: context.parameters.AddTagsOnBlur.raw || false,
			duplicates: false,
			editTags: false,
			backspace: true,
			placeholder: context.parameters.Placeholder.raw || this._resx.getString('AddTagKey'),
			dropdown: {
				enabled: context.parameters.AutocompleteEnabled.raw || true,
				maxItems: context.parameters.AutocompleteMaxItems.raw || 10,
				highlightFirst: context.parameters.AutocompleteHighlightFirst.raw || false,
				closeOnSelect: context.parameters.AutocompleteCloseOnSelect.raw || true,
				position: context.parameters.AutocompletePosition.raw || 'all',
				fuzzySearch: context.parameters.AutocompleteEnableFuzzySearch.raw || true
			}
		});

		// if the application insights instrumentation key were set, it will init the Application Insights Service Proxy
		const appInsightsInstrumentationKey: string = context.parameters.AppInsightsInstrumentationKey.raw || '';
		if (appInsightsInstrumentationKey != '') {
			this._appInsights = new ApplicationInsights({
				config: {
					instrumentationKey: appInsightsInstrumentationKey,
				}
			});
			this._appInsights.loadAppInsights();
		}

		// initiliazes the Tagify (https://github.com/yairEO/tagify) control by yairEO (https://github.com/yairEO)
		this._tagify
			.on('input', (e) => { this.onInput(e) })
			.on('remove', (e) => { this.onRemove(e) })
			.on('add', (e) => { if (this._initialized) this.onAdd(e) })
			.on('edit:updated', (e) => { this.onUpdate(e) });


		// tries to find the <tag> element
		this.findTagElement();

		// loads all initial Tag values from the many to many relationship
		this.InitValues();

	}


	/**
	 * @see https://dynamicsninja.blog/2019/11/25/is-your-pcf-control-read-only/ 
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		// Add code to update control view
		this._isDisabled = context.mode.isControlDisabled;
		this._isVisible = context.mode.isVisible;
		if (this._tagEl) {
			// no need to test _isVisible as per https://www.itaintboring.com/dynamics-crm/pcf-components-and-setvisible-setdisabled/
			if (this._isDisabled) {
				// put the tag on readonly mode to prevent user interaction.
				this._tagEl.setAttribute('readonly', 'readonly');
			} else {
				this._tagEl.removeAttribute('readonly');
			}
		}
		//this.applySettings(context);
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		return {};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		// Add code to cleanup control if necessary
		if(this._tagify){
			this._currentState = '';
			this._tagify.destroy();
		}
	}


	/**
	 * Load all initial values (tags) based on the many to many relationship of the current entity and the Tag entity.
	 */
	private async InitValues() {
		try {
			let tags: Array<Tag> = await this._tagManagerService.GetAllByRecordId();
			this._tagify.addTags(tags);
		} catch (err) {
			this.logError(err);
			if (err instanceof CustomError) {
				this.displayUserMessage(err.message);
			} else {
				this.displayUserMessage(this._resx.getString('DefaultErrorMessageKey'));
			}
		} finally {
			// now the onAdd method can be triggered.
			this._initialized = true;
		}
	}

	/**
	 * Load all tags from server while the user is typing something on the input.
	 * This method will query all Tags and all Translated Tags based on the user input.
	 * @param e the Tagify event data info.
	 */
	private async onInput(e: any) {
		try {
			let inputValue: string = (e.detail.value || "").replace(/\s/g, "");
			if (this._currentState != 'error' && (inputValue.length >= this._searchActivationTreshold)) {
				// reset current whitelist
				this._tagify.settings.whitelist.length = 0;

				// show loading animation and hide the suggestions dropdown
				this._tagify.loading(true).dropdown.hide.call(this._tagify)

				// https://developer.mozilla.org/en-US/docs/Web/API/AbortController/abort
				let controller: AbortController = new AbortController();
				if (controller) {
					controller.abort();
				}

				let tags: Array<Tag> = await this._tagManagerService.SearchByName(inputValue);

				// replace tagify "whitelist" array values with new values 
				this._tagify.settings.whitelist = tags;

				// render the suggestions dropdown
				this._tagify.loading(false).dropdown.show.call(this._tagify, e.detail.value);
			}
		} catch (err) {
			this.logError(err);
			if (err instanceof CustomError) {
				this.displayUserMessage(err.message);
			} else {
				this.displayUserMessage(this._resx.getString('DefaultErrorMessageKey'));
			}
		} finally {
			if (this._tagify) {
				this._tagify.loading(false);
			}
		}
	}

	/**
	 * The onAdd event handler that will be triggered only if the control was initiliazed, which means that all Tags were loaded. 
	 * If the state of this component is equals to 'error', or if the Tag is a 'temp' Tag, the event will trigger, but nothing will happen. 
	 * If any there was any error during the Tag creation, the Tag will be flagged as 'temp' since it shouldn't invoke the Tag Manager Service.
	 * @param e the Tagify event data info
	 */
	private async onAdd(e: any) {

		let addedTagInputs: Tag = e.detail.data;
		try {

			if (this._currentState != 'error') {
				if (!this._tagEl) {
					this.findTagElement();
				}

				if (!this._tagEl) {
					throw new CustomError(this._resx.getString('ErrTagCantBeAddedKey').format({ tagName: addedTagInputs.value }));
				}

				// The first validation covers any user Tag's selection from the autocomplete list. The second validation covers the situation where the user wants to create a new Tag.
				if ((addedTagInputs.id != null && addedTagInputs.persisted) || (addedTagInputs.id == null && addedTagInputs.value != null)) {

					// put the tag on readonly mode to prevent user interaction.
					this._tagEl.setAttribute('readonly', 'readonly');

					let tag: Tag | null = await this._tagManagerService.Add(addedTagInputs.value, addedTagInputs.id);

					if (tag && !addedTagInputs.id) {
						this._tagify.replaceTag(e.detail.tag, tag); // replace the content with the added tag info;
					}

					if (!tag || (tag && !tag.id)) {
						throw new CustomError(this._resx.getString('ErrTagWasNotAddedKey').format({ tagName: addedTagInputs.value }));
					}
				}
			}
		} catch (err) {
			this.logError(err);
			if (err instanceof CustomError) {
				this.displayUserMessage(err.message);
			} else {
				this.displayUserMessage(this._resx.getString('DefaultErrorMessageKey'));
			}
		} finally {
			// "release" the readonly mode if the current state isn't error, which means that the Tag was added.
			if (this._tagEl && this._currentState != 'error') {
				this._tagEl.removeAttribute('readonly');
			} else if (this._currentState == 'error' && (addedTagInputs && !(<any>addedTagInputs).temp)) {
				// replace the content with the persisted flag set and this will not trigger the onRemove event	
				this._tagify.replaceTag(e.detail.tag, { value: e.detail.data.value, persisted: false });
				this._tagify.removeTag(e.detail.data.value);
			}
		}
	}

	/**
	 * The onRemove event Handler. If the state is equals to error, or if the Tag wasn't persisted (saved on CDS), the event will trigger, but nothing will happen.
	 * If the Tag was not removed, this event handler will automatically add the same Tag again and a message will be displayed to the user.
	 * @param e the Tagify event data info
	 */
	private async onRemove(e: any) {

		let removedTagInputs: Tag = e.detail.data;
		try {
			if (this._currentState != 'error' && (removedTagInputs.id && removedTagInputs.persisted)) {
				let tagRemoved: Boolean = await this._tagManagerService.Remove(removedTagInputs.id);
				if (!tagRemoved) {
					throw new CustomError(this._resx.getString('ErrTagWasNotRemovedKey').format({ tagName: removedTagInputs.value }));
				}
			}
		} catch (err) {
			this.logError(err);
			if (err instanceof CustomError) {
				this.displayUserMessage(err.message);
			} else {
				this.displayUserMessage(this._resx.getString('DefaultErrorMessageKey'));
			}
		} finally {
			if (this._currentState == 'error' && (removedTagInputs && removedTagInputs.persisted)) {
				// by doing the addTags first with the flag 'temp' and the replaceTag latter, it will prevent the onAdd event to parse the inputs;
				let addedTag: any = this._tagify.addTags([{ value: removedTagInputs.value, temp: true }]);
				this._tagify.replaceTag(addedTag[0], removedTagInputs);
			}
		}
	}

	/**
	 * This should update the main tag or the translation...
	 * This method will be implemented soon. Lot of validations to cover here since this component is a Multi Language component.
	 * @param e the Tagify event data info
	 */
	private async onUpdate(e: any) {
		let updatedTagInputs: Tag = e.detail.data;
		try {
			this._tagManagerService.Update(<string>updatedTagInputs.id, '');
		} catch (err) {
			this.logError(err);
			if (err instanceof CustomError) {
				this.displayUserMessage(err.message);
			} else {
				this.displayUserMessage(this._resx.getString('DefaultErrorMessageKey'));
			}
		}
	}

	/**
	 * (not used) Apply the general settings to the tagify input everytime the property bag changes.
	 * This method is still not working as expected, that's the reason it has being called.
	 * @param context 
	 */
	private applySettings(context: ComponentFramework.Context<IInputs>): void {
		let settings: any = {
			maxTags: context.parameters.AutocompleteMaxItems.raw || 10,
			enforceWhitelist: context.parameters.EnforceWhitelist.raw || false,
			addTagOnBlur: context.parameters.AddTagsOnBlur.raw || true,
			duplicates: false,
			editTags: false,
			backspace: true,
			placeholder: context.parameters.Placeholder.raw || 'Add tag',
			dropdown: {
				enabled: context.parameters.AutocompleteEnabled.raw || true,
				maxItems: context.parameters.AutocompleteMaxItems.raw || 10,
				highlightFirst: context.parameters.AutocompleteHighlightFirst.raw || false,
				closeOnSelect: context.parameters.AutocompleteCloseOnSelect.raw || true,
				position: context.parameters.AutocompletePosition.raw || 'input',
				fuzzySearch: context.parameters.AutocompleteEnableFuzzySearch.raw || true
			}
		};
		this._tagify.settings = { ...this._tagify.settings, ...settings };
	}

	/**
	 * Log any kind of error on the Azure Application Insights Service or on the Console. 
	 * In order to use the Application Insights Service, the property AppInsightsInstrumentationKey can't be empty.
	 * @param err the Error instance that will be logged.
	 */
	private logError(err: Error) {
		try {
			if (this._appInsights) {
				let exception: IExceptionTelemetry = { exception: err };
				this._appInsights.trackException(exception);
			} else console.error(err.message, err.stack);
		} catch (e) {
			console.error(err.message, err.stack);
		}
	}

	/**
	 * Display any kind of user message, mostly used to display error messages right below the input.
	 * @param message the message that will be displayed to the user
	 * @param isError indicates that an error happened and the notification will be displayed in red with a X icon as stated: https://docs.microsoft.com/en-us/powerapps/developer/model-driven-apps/clientapi/reference/controls/setnotification 
	 */
	private displayUserMessage(message: string, isError: boolean = true, actions?: Array<any>) {
		const notificationId: string = '_' + Math.random().toString(36).substr(2, 9);
		const clientApi: any = (<any>this._context.utils);
		if (isError) {
			if (this._tagEl != null) {
				this._tagEl.setAttribute('readonly', 'readonly');
				this._currentState = 'error';
			}
			clientApi.setNotification(message + ' ' + this._resx.getString('RefreshPageTryAgainKey'), notificationId);
		} else {
			clientApi.addNotificaiton({
				messages: [message],
				notificationLevel: 'RECOMMENDATION',
				uniqueId: notificationId,
				actions: actions
			});
		}

		window.setTimeout(() => {
			clientApi.clearNotification(notificationId);
			// "release" the readonly mode
			if (this._tagEl) {
				this._tagEl.removeAttribute('readonly');
				this._currentState = '';
			}
		}, 5000);
	}

	/**
	 * Auxiliar method. Tries to find the <tag> element on the container root element.
	 * This element will be used to set attributes such as readonly, with, visibility etc.
	 */
	private findTagElement() {
		if (!this._tagEl) {
			let htmlElements: HTMLCollectionBase = this._container.getElementsByTagName('tags');
			if (htmlElements != null && htmlElements.length > 0) {
				this._tagEl = htmlElements[0] as HTMLElement;
			}
		}
	}
}