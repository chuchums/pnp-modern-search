import * as React from 'react';
import {
    IPropertyPaneGroup,
    IPropertyPanePage,
    IPropertyPaneDropdownOption,
    PropertyPaneChoiceGroup,
    PropertyPaneDropdown,
    PropertyPaneLabel,
    PropertyPaneTextField,
    PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { IComboBoxOption } from '@fluentui/react/lib/ComboBox';
import { IAsyncComboProps } from '../../../controls/PropertyPaneAsyncCombo/components/IAsyncComboProps';
import { IAdsEnvoiRedactionSettings } from '../../../models/common/IAdsEnvoiRedactionSettings';

const ROOT = 'adsEnvoiRedaction';
const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const USER_TYPES = ['User', 'UserMulti'];
/** Propriétés nécessaires au lancement du flux */
const REQUIRED_PROPERTIES = ['ListItemID', 'Path', 'Filename'];

interface IListField { key: string; text: string; type: string; }

/**
 * ADS GED – onglet « Envoi en rédaction » de la webpart Résultats.
 * L'instance est conservée par la webpart pour garder en cache les listes et colonnes chargées.
 */
export class AdsEnvoiRedactionPropertyPaneBuilder {

    private _lists: IPropertyPaneDropdownOption[] = [];
    private _fields: IListField[] = [];
    private _loadedSiteUrl: string = null;
    private _loadedListId: string = null;
    private _loading = false;
    private _loadError: string = null;

    constructor(
        private spHttpClient: SPHttpClient,
        private currentWebUrl: string,
        private refreshPropertyPane: () => void
    ) { }

    /** Charge (si nécessaire) les listes du site et les colonnes de la liste sélectionnée, puis rafraîchit le panneau. */
    public async ensureLoaded(settings: IAdsEnvoiRedactionSettings): Promise<void> {
        const siteUrl = this._siteUrl(settings);
        const listId = settings?.configListId || null;
        if (this._loading) return;
        if (siteUrl === this._loadedSiteUrl && listId === this._loadedListId) return;

        this._loading = true;
        this._loadError = null;
        try {
            if (siteUrl !== this._loadedSiteUrl) {
                this._lists = await this._loadLists(siteUrl);
                this._loadedSiteUrl = siteUrl;
                this._fields = [];
                this._loadedListId = null;
            }
            if (listId && listId !== this._loadedListId) {
                this._fields = await this._loadFields(siteUrl, listId);
            } else if (!listId) {
                this._fields = [];
            }
            this._loadedListId = listId;
        } catch (error) {
            this._loadError = `${error && error.message ? error.message : error}`;
            this._loadedSiteUrl = siteUrl; // évite une boucle de rechargement
            this._loadedListId = listId;
        } finally {
            this._loading = false;
            this.refreshPropertyPane();
        }
    }

    /** Force le rechargement au prochain affichage (ex : changement de site). */
    public invalidate(): void {
        this._loadedSiteUrl = null;
        this._loadedListId = null;
    }

    public buildPage(
        settings: IAdsEnvoiRedactionSettings,
        selectedProperties: string[],
        propertyFieldCollectionData: any,
        customCollectionFieldType: any,
        asyncComboComponent: any
    ): IPropertyPanePage {

        settings = settings || ({} as IAdsEnvoiRedactionSettings);
        const enabled = !!settings.enabled;

        const managedPropertyOptions: IComboBoxOption[] = (selectedProperties || []).map(p => ({ key: p, text: p }));
        const listFieldOptions: IDropdownOption[] = this._fields.map(f => ({ key: f.key, text: `${f.text} (${f.key})` }));
        const userFieldOptions: IPropertyPaneDropdownOption[] = this._fields
            .filter(f => USER_TYPES.indexOf(f.type) > -1)
            .map(f => ({ key: f.key, text: `${f.text} (${f.key})` }));

        // ---- Bouton
        const buttonGroup: IPropertyPaneGroup = {
            groupName: 'Bouton',
            isCollapsed: false,
            groupFields: [
                PropertyPaneToggle(`${ROOT}.enabled`, {
                    label: 'Afficher le bouton dans les résultats',
                    onText: 'Activé',
                    offText: 'Désactivé',
                    checked: enabled
                }),
                PropertyPaneTextField(`${ROOT}.buttonLabel`, {
                    label: 'Libellé du bouton',
                    placeholder: 'Envoi en rédaction',
                    value: settings.buttonLabel,
                    disabled: !enabled
                }),
                PropertyPaneTextField(`${ROOT}.confirmText`, {
                    label: 'Message de confirmation (optionnel)',
                    description: '{fileName} est remplacé par le nom du document.',
                    placeholder: 'Envoyer « {fileName} » en rédaction ?',
                    multiline: true,
                    rows: 2,
                    value: settings.confirmText,
                    disabled: !enabled
                })
            ]
        };

        // ---- Flux
        const flowGroup: IPropertyPaneGroup = {
            groupName: 'Flux Power Automate',
            isCollapsed: !enabled,
            groupFields: [
                PropertyPaneTextField(`${ROOT}.flowId`, {
                    label: 'Id du flux',
                    description: 'GUID du flux « Pour un fichier sélectionné », ou format v1/<envId>/<flowId>.',
                    placeholder: '26c31753-a040-acc9-344d-6a9e65bae278',
                    value: settings.flowId,
                    disabled: !enabled,
                    onGetErrorMessage: (value: string) => this._validateFlowId(value, settings)
                }),
                PropertyPaneTextField(`${ROOT}.environmentId`, {
                    label: 'Id de l\'environnement (optionnel)',
                    description: 'Vide = environnement par défaut du tenant.',
                    value: settings.environmentId,
                    disabled: !enabled
                }),
                PropertyPaneTextField(`${ROOT}.flowTriggerUri`, {
                    label: 'URL de déclenchement « logicflows » (optionnel)',
                    description: 'Ex : https://europe-001.azure-apim.net/apim/logicflows/<ID>/triggers/manual/run. Visible dans l\'onglet Réseau lors d\'un « Tester » du flux. Vide = découverte automatique.',
                    value: settings.flowTriggerUri,
                    disabled: !enabled
                }),
                PropertyPaneTextField(`${ROOT}.flowExtraBody`, {
                    label: 'Données complémentaires envoyées au flux (JSON, optionnel)',
                    description: 'Fusionné dans le corps de la requête. Ex : {"text": "{userEmail}"} pour une entrée du déclencheur, ou {"entity": {"Champ": "{ID}"}}. Variables : {ID} {itemUrl} {fileName} {userEmail} {userName}.',
                    multiline: true,
                    rows: 4,
                    value: settings.flowExtraBody,
                    disabled: !enabled,
                    onGetErrorMessage: (value: string) => {
                        const v = (value || '').trim();
                        if (!v) return '';
                        try { JSON.parse(v.replace(/\{(ID|itemUrl|fileName|userEmail|userName)\}/g, 'x')); return ''; } catch { return 'JSON invalide.'; }
                    }
                }),
                PropertyPaneTextField(`${ROOT}.flowUrl`, {
                    label: 'URL de déclencheur HTTP (optionnel)',
                    description: 'Si renseignée, le flux est appelé via cette URL au lieu de son Id.',
                    value: settings.flowUrl,
                    disabled: !enabled
                })
            ]
        };

        // ---- Liste de paramétrage
        const listFields: any[] = [
            PropertyPaneTextField(`${ROOT}.configSiteUrl`, {
                label: 'Site de la liste de paramétrage',
                description: 'Vide = site courant.',
                placeholder: this.currentWebUrl,
                value: settings.configSiteUrl,
                disabled: !enabled
            }),
            PropertyPaneDropdown(`${ROOT}.configListId`, {
                label: 'Liste de paramétrage',
                options: this._lists,
                selectedKey: settings.configListId,
                disabled: !enabled || this._lists.length === 0
            }),
            PropertyPaneDropdown(`${ROOT}.redacteursField`, {
                label: 'Colonne des rédacteurs (Personne)',
                options: userFieldOptions,
                selectedKey: settings.redacteursField,
                disabled: !enabled || userFieldOptions.length === 0
            }),
            PropertyPaneChoiceGroup(`${ROOT}.matchMode`, {
                label: 'Le bouton s\'affiche si',
                options: [
                    { key: 'any', text: 'au moins une colonne mappée correspond (ex : Service OU Processus)', checked: settings.matchMode !== 'all' },
                    { key: 'all', text: 'toutes les colonnes renseignées de la ligne correspondent', checked: settings.matchMode === 'all' }
                ]
            })
        ];

        if (this._loading) {
            listFields.push(PropertyPaneLabel(`${ROOT}.loading`, { text: 'Chargement des listes et colonnes…' }));
        }
        if (this._loadError) {
            listFields.push(PropertyPaneLabel(`${ROOT}.error`, { text: `⚠ ${this._loadError}` }));
        }

        if (propertyFieldCollectionData && customCollectionFieldType) {
            listFields.push(
                propertyFieldCollectionData(`${ROOT}.mappings`, {
                    key: `${ROOT}.mappings`,
                    label: 'Correspondance colonnes ↔ propriétés gérées',
                    panelHeader: 'Correspondance colonnes de la liste ↔ propriétés gérées de recherche',
                    panelDescription: 'Pour chaque colonne de la liste de paramétrage, indiquez la propriété gérée qui porte la même information sur le document. La propriété gérée doit figurer dans les « Propriétés sélectionnées » de la source de données.',
                    manageBtnLabel: 'Configurer le mapping',
                    value: settings.mappings,
                    disabled: !enabled || !settings.configListId,
                    enableSorting: false,
                    fields: [
                        {
                            id: 'listField',
                            title: 'Colonne de la liste',
                            type: customCollectionFieldType.custom,
                            required: true,
                            onCustomRender: (field, value, onUpdate, item, itemId) => {
                                return React.createElement(Dropdown, {
                                    key: `lf-${itemId}`,
                                    options: listFieldOptions,
                                    selectedKey: value || null,
                                    placeholder: 'Choisir une colonne',
                                    onChange: (_e, option: IDropdownOption) => onUpdate(field.id, option ? option.key : '')
                                });
                            }
                        },
                        {
                            id: 'managedProperty',
                            title: 'Propriété gérée',
                            type: customCollectionFieldType.custom,
                            required: true,
                            onCustomRender: (field, value, onUpdate, item, itemId) => {
                                return React.createElement(asyncComboComponent, {
                                    key: `mp-${itemId}`,
                                    allowFreeform: true,
                                    availableOptions: managedPropertyOptions,
                                    placeholder: 'ex : owstaxIdService',
                                    textDisplayValue: value || '',
                                    defaultSelectedKey: value || '',
                                    onLoadOptions: () => Promise.resolve(managedPropertyOptions),
                                    onUpdateOptions: () => { /* noop */ },
                                    onUpdate: (option: IComboBoxOption) => onUpdate(field.id, option ? option.key : '')
                                } as IAsyncComboProps);
                            }
                        }
                    ]
                })
            );
        }

        // Propriétés gérées manquantes dans la source de données
        const selectedLower = (selectedProperties || []).map(p => p.toLowerCase());
        const needed = REQUIRED_PROPERTIES.concat((settings.mappings || []).map(m => m && m.managedProperty).filter(Boolean));
        const missing = needed.filter((p, i) => needed.indexOf(p) === i && selectedLower.indexOf(p.toLowerCase()) === -1);
        if (enabled && missing.length > 0) {
            listFields.push(PropertyPaneLabel(`${ROOT}.missing`, {
                text: `⚠ À ajouter aux propriétés sélectionnées de la source de données : ${missing.join(', ')}`
            }));
        }

        const listGroup: IPropertyPaneGroup = {
            groupName: 'Liste de paramétrage et correspondances',
            isCollapsed: !enabled,
            groupFields: listFields
        };

        // ---- Intégration dans le template
        const templateGroup: IPropertyPaneGroup = {
            groupName: 'Intégration dans le template',
            isCollapsed: true,
            groupFields: [
                PropertyPaneLabel(`${ROOT}.snippet`, {
                    text: 'Dans la boucle {{#each data.items as |item|}} du template, ajoutez : ' +
                        '<ads-envoi-redaction data-config="{{JSONstringify @root.properties.adsEnvoiRedaction}}" ' +
                        'data-item="{{JSONstringify item}}" data-theme-variant="{{JSONstringify @root.theme}}"></ads-envoi-redaction>'
                })
            ]
        };

        return {
            header: { description: 'Bouton « Envoi en rédaction » : affiché uniquement aux rédacteurs définis dans la liste de paramétrage.' },
            displayGroupsAsAccordion: true,
            groups: [buttonGroup, flowGroup, listGroup, templateGroup]
        };
    }

    private _siteUrl(settings: IAdsEnvoiRedactionSettings): string {
        return ((settings && settings.configSiteUrl) || this.currentWebUrl || '').trim().replace(/\/$/, '');
    }

    private _validateFlowId(value: string, settings: IAdsEnvoiRedactionSettings): string {
        if (!settings || !settings.enabled) return '';
        const v = (value || '').trim();
        if (!v) return (settings.flowUrl || '').trim() ? '' : 'Renseignez l\'Id du flux (ou une URL de déclencheur HTTP).';
        const parts = v.split('/');
        const id = parts.length === 3 && parts[0].toLowerCase() === 'v1' ? parts[2] : v;
        return GUID_RE.test(id) ? '' : 'Format attendu : GUID ou v1/<envId>/<flowId>.';
    }

    private async _loadLists(siteUrl: string): Promise<IPropertyPaneDropdownOption[]> {
        const res: SPHttpClientResponse = await this.spHttpClient.get(
            `${siteUrl}/_api/web/lists?$select=Id,Title&$filter=Hidden eq false and BaseType eq 0&$orderby=Title`,
            SPHttpClient.configurations.v1,
            { headers: { Accept: 'application/json;odata=nometadata' } }
        );
        if (!res.ok) throw new Error(`Impossible de lire les listes du site (${res.status})`);
        const data = await res.json();
        return (data.value || []).map((l: any) => ({ key: l.Id, text: l.Title }));
    }

    private async _loadFields(siteUrl: string, listId: string): Promise<IListField[]> {
        const res: SPHttpClientResponse = await this.spHttpClient.get(
            `${siteUrl}/_api/web/lists(guid'${listId}')/fields?$select=InternalName,Title,TypeAsString&$filter=Hidden eq false&$orderby=Title`,
            SPHttpClient.configurations.v1,
            { headers: { Accept: 'application/json;odata=nometadata' } }
        );
        if (!res.ok) throw new Error(`Impossible de lire les colonnes de la liste (${res.status})`);
        const data = await res.json();
        const technical = ['ContentType', 'Attachments', 'Edit', 'DocIcon', 'LinkTitleNoMenu', 'LinkTitle', 'ItemChildCount', 'FolderChildCount', '_ComplianceFlags', '_ComplianceTag', '_ComplianceTagWrittenTime', '_ComplianceTagUserId', '_IsRecord', 'AppAuthor', 'AppEditor', '_UIVersionString', 'ComplianceAssetId'];
        return (data.value || [])
            .filter((f: any) => technical.indexOf(f.InternalName) === -1)
            .map((f: any) => ({ key: f.InternalName, text: f.Title, type: f.TypeAsString }));
    }
}
