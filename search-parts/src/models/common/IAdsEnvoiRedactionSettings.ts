/**
 * ADS GED – paramétrage du bouton « Envoi en rédaction » (onglet dédié de la webpart Résultats).
 */
export interface IAdsColumnMapping {
    /** Nom interne de la colonne dans la liste de paramétrage (ex: Service) */
    listField: string;
    /** Propriété gérée de recherche correspondante (ex: owstaxIdService, RefinableString100) */
    managedProperty: string;
}

export interface IAdsEnvoiRedactionSettings {
    /** Active l'affichage du bouton */
    enabled: boolean;
    /** Libellé du bouton */
    buttonLabel: string;
    /** Texte de confirmation ({fileName} est remplacé par le nom du document) */
    confirmText?: string;

    /** Id du flux Power Automate (GUID ou "v1/<envId>/<flowId>") */
    flowId: string;
    /** Id de l'environnement Power Platform (vide = environnement par défaut) */
    environmentId?: string;
    /** Hôte de l'API Flow (vide = https://api.flow.microsoft.com) */
    flowApiHost?: string;
    /** Alternative : URL d'un déclencheur HTTP. Si renseignée, remplace l'appel à l'API Flow */
    flowUrl?: string;
    /**
     * URL « logicflows » du déclencheur (https://<région>.azure-apim.net/apim/logicflows/<id>/triggers/manual/run).
     * Vide = découverte automatique depuis l'Id du flux.
     */
    flowTriggerUri?: string;
    /**
     * JSON fusionné dans le corps envoyé au flux (entrées manuelles du déclencheur, champs requis…).
     * Une clé "entity" complète rows[0].entity. Variables : {ID} {itemUrl} {fileName} {userEmail} {userName}
     */
    flowExtraBody?: string;

    /** Site hébergeant la liste de paramétrage (vide = site courant) */
    configSiteUrl?: string;
    /** Id (GUID) de la liste de paramétrage */
    configListId: string;
    /** Nom interne de la colonne Personne « Rédacteurs » */
    redacteursField: string;
    /** Correspondances colonne de liste ↔ propriété gérée */
    mappings: IAdsColumnMapping[];
    /** 'any' : une colonne mappée suffit (Service OU Processus) ; 'all' : toutes les colonnes renseignées doivent correspondre */
    matchMode: 'any' | 'all';
}

export const DEFAULT_ADS_ENVOI_REDACTION_SETTINGS: IAdsEnvoiRedactionSettings = {
    enabled: false,
    buttonLabel: 'Envoi en rédaction',
    confirmText: '',
    flowId: '',
    environmentId: '',
    flowApiHost: '',
    flowUrl: '',
    flowTriggerUri: '',
    flowExtraBody: '',
    configSiteUrl: '',
    configListId: '',
    redacteursField: '',
    mappings: [],
    matchMode: 'any'
};
