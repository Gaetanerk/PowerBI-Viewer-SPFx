import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  type IPropertyPaneDropdownOption,
  type IPropertyPaneField,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import strings from './loc/strings';
import PowerBiViewer, { IPowerBiViewerProps } from './components/PowerBiViewer';
import { IPowerBiService, IReportSummary } from './models/IPowerBiModels';
import { createAadTransport, normalizeApiBaseUrl, PowerBiApiService } from './services/PowerBiApiService';
import { API_RESOURCE, SIMULATION_DEFAULT_URL } from './services/apiConfig';

/** Propriétés enregistrées dans la page : uniquement de la configuration, jamais de jeton ni de secret. */
export interface IPowerBiViewerWebPartProps {
  apiBaseUrl: string;
  reportId: string;
  autoHeight: boolean;
  customHeight: number;
  /** Ignoré dans le package de production (cf. _createService). */
  mockMode?: boolean;
}

type CatalogState =
  | { status: 'idle' }
  | { status: 'loading' }
  | { status: 'loaded'; reports: IReportSummary[] }
  | { status: 'unavailable' };

export default class PowerBiViewerWebPart extends BaseClientSideWebPart<IPowerBiViewerWebPartProps> {
  private _serviceKey: string | undefined;
  private _servicePromise: Promise<IPowerBiService | undefined> = Promise.resolve(undefined);
  private _catalog: CatalogState = { status: 'idle' };
  private _catalogKey: string | undefined;

  public render(): void {
    const servicePromise = this._getService();
    servicePromise
      .then(service => {
        // Ignore un rendu devenu obsolète si la configuration a changé entre-temps.
        if (servicePromise !== this._servicePromise) {
          return;
        }
        const props: IPowerBiViewerProps = {
          service,
          reportId: (this.properties.reportId || '').trim(),
          fixedHeight: this.properties.autoHeight === false ? this.properties.customHeight : undefined,
          isEditMode: this.displayMode === DisplayMode.Edit,
          simulationUrl: this._getSimulationUrl()
        };
        ReactDom.render(React.createElement(PowerBiViewer, props), this.domElement);
      })
      .catch(error => console.error('[PowerBiViewer] Initialisation du service impossible', error));
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /** URL du simulateur local lorsque le mode simulation est actif (builds de développement uniquement). */
  private _getSimulationUrl(): string | undefined {
    return DEBUG && this.properties.mockMode
      ? normalizeApiBaseUrl(this.properties.apiBaseUrl) || SIMULATION_DEFAULT_URL
      : undefined;
  }

  /** Un seul service par configuration, pour ne pas relancer le chargement du rapport à chaque rendu. */
  private _getService(): Promise<IPowerBiService | undefined> {
    const baseUrl = normalizeApiBaseUrl(this.properties.apiBaseUrl);
    const simulationUrl = this._getSimulationUrl();
    const key = simulationUrl ? `simulation:${simulationUrl}` : baseUrl || '';
    if (key !== this._serviceKey) {
      this._serviceKey = key;
      this._servicePromise = this._createService(baseUrl, simulationUrl);
    }
    return this._servicePromise;
  }

  private _createService(baseUrl: string | undefined, simulationUrl: string | undefined): Promise<IPowerBiService | undefined> {
    // DEBUG vaut false dans un build --production : webpack supprime alors cette branche et le transport de
    // simulation n'est pas inclus dans le package. Ne pas transformer cette méthode en "async", sinon la
    // branche devient un état de machine à états et n'est plus éliminée.
    if (DEBUG && simulationUrl) {
      return import(/* webpackChunkName: 'simulation-transport' */ './services/SimulationTransport')
        .then(module => new PowerBiApiService(module.createSimulationTransport(), simulationUrl));
    }
    if (!baseUrl) {
      return Promise.resolve(undefined);
    }
    // Jeton délégué obtenu silencieusement pour l'utilisateur connecté à SharePoint, audience = API d'intégration.
    return this.context.aadHttpClientFactory
      .getClient(API_RESOURCE)
      .then(client => new PowerBiApiService(createAadTransport(client), baseUrl));
  }

  protected onPropertyPaneConfigurationStart(): void {
    this._loadCatalog();
  }

  protected onPropertyPaneFieldChanged(propertyPath: string): void {
    if (propertyPath === 'apiBaseUrl' || propertyPath === 'mockMode') {
      this._loadCatalog();
    }
  }

  /** Charge la liste des rapports si l'API l'expose ; sinon le panneau propose une saisie manuelle. */
  private _loadCatalog(): void {
    const servicePromise = this._getService();
    const key = this._serviceKey;
    if (key === this._catalogKey) {
      return;
    }
    this._catalogKey = key;
    this._catalog = { status: key ? 'loading' : 'idle' };
    this.context.propertyPane.refresh();
    if (!key) {
      return;
    }

    servicePromise
      .then(service => (service ? service.getReports() : []))
      .then(reports => {
        if (key === this._catalogKey) {
          this._catalog = { status: 'loaded', reports };
          this.context.propertyPane.refresh();
        }
      })
      .catch(() => {
        if (key === this._catalogKey) {
          this._catalog = { status: 'unavailable' };
          this._catalogKey = undefined; // Nouvelle tentative à la prochaine ouverture du panneau.
          this.context.propertyPane.refresh();
        }
      });
  }

  private _getReportField(): IPropertyPaneField<unknown> {
    const catalog = this._catalog;
    const reportId = this.properties.reportId;

    if (catalog.status === 'loading') {
      return PropertyPaneDropdown('reportId', { label: strings.propertyPane.reportLoading, options: [], disabled: true });
    }
    if (catalog.status === 'loaded' && catalog.reports.length > 0) {
      const options: IPropertyPaneDropdownOption[] = catalog.reports.map(report => ({ key: report.id, text: report.name }));
      if (reportId && !catalog.reports.some(report => report.id === reportId)) {
        options.unshift({ key: reportId, text: `${reportId} ${strings.propertyPane.reportNotInCatalog}` });
      }
      return PropertyPaneDropdown('reportId', { label: strings.propertyPane.reportLabel, options, selectedKey: reportId });
    }

    let description = strings.propertyPane.reportManualDescription;
    if (catalog.status === 'unavailable') {
      description = strings.propertyPane.reportCatalogUnavailable;
    } else if (catalog.status === 'loaded') {
      description = strings.propertyPane.reportCatalogEmpty;
    }
    return PropertyPaneTextField('reportId', {
      label: strings.propertyPane.reportLabel,
      description,
      // La validation différée évite d'appeler l'API à chaque frappe.
      onGetErrorMessage: (value: string) => (value && !value.trim() ? strings.propertyPane.reportRequired : ''),
      deferredValidationTime: 800
    });
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const isMock = DEBUG && !!this.properties.mockMode;
    const connectionFields: IPropertyPaneField<unknown>[] = [];

    if (DEBUG) {
      connectionFields.push(PropertyPaneToggle('mockMode', {
        label: strings.propertyPane.mockModeLabel,
        onText: strings.propertyPane.mockModeOn,
        offText: strings.propertyPane.mockModeOff
      }));
    }
    connectionFields.push(
      PropertyPaneTextField('apiBaseUrl', {
        label: strings.propertyPane.apiBaseUrlLabel,
        description: isMock ? strings.propertyPane.apiBaseUrlSimulationDescription : strings.propertyPane.apiBaseUrlDescription,
        placeholder: isMock ? SIMULATION_DEFAULT_URL : 'https://',
        onGetErrorMessage: (value: string) => (!value || normalizeApiBaseUrl(value) ? '' : strings.propertyPane.apiBaseUrlInvalid),
        deferredValidationTime: 800
      }),
      this._getReportField()
    );

    const displayFields: IPropertyPaneField<unknown>[] = [
      PropertyPaneToggle('autoHeight', {
        label: strings.propertyPane.autoHeightLabel,
        onText: strings.propertyPane.autoHeightOn,
        offText: strings.propertyPane.autoHeightOff
      })
    ];
    if (this.properties.autoHeight === false) {
      displayFields.push(PropertyPaneSlider('customHeight', {
        label: strings.propertyPane.customHeightLabel,
        min: 300,
        max: 2000,
        step: 50,
        showValue: true
      }));
    }

    return {
      pages: [
        {
          header: { description: strings.propertyPane.description },
          groups: [
            { groupName: strings.propertyPane.connectionGroup, groupFields: connectionFields },
            { groupName: strings.propertyPane.displayGroup, groupFields: displayFields }
          ]
        }
      ]
    };
  }
}
