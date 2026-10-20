import * as React from "react";
import * as ReactDOM from "react-dom";
import { BaseWebComponent } from "@pnp/modern-search-extensibility";
import {
  ActionButton, IconButton, ITheme, Link, MessageBar, MessageBarType, Panel, PanelType,
  Spinner, SpinnerSize, TooltipHost, getTheme
} from "@fluentui/react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { PageContext } from "@microsoft/sp-page-context";
import { Log } from "@microsoft/sp-core-library";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

/**
 * ADS GED – bouton « Détail » (icône) : affiche les propriétés SharePoint du document
 * dans un panneau latéral, sans quitter la page de recherche.
 *
 * Dans le template Handlebars :
 *
 *   <ads-detail-document
 *       data-item="{{JSONstringify item}}"
 *       data-theme-variant="{{JSONstringify @root.theme}}">
 *   </ads-detail-document>
 *
 * Attributs facultatifs :
 *   data-fields="Title,NatureDocument,Service"   noms internes à afficher, dans cet ordre
 *                                                (par défaut : toutes les colonnes visibles du type de contenu)
 *   data-exclude-fields="Title"                  noms internes à masquer
 *   data-label="Détail"                          infobulle / titre accessible de l'icône
 *   data-icon="Info"                             icône Fluent UI
 *
 * Propriétés gérées utilisées : SPWebUrl, NormListID (ou ListID), ListItemID ; à défaut Path.
 * Les valeurs sont lues en direct (REST, droits de l'utilisateur), au premier clic, puis gardées en cache.
 */

const LOG_SOURCE = "AdsDetailDocument";

export interface IAdsDetailDocumentProps {
  item?: { [key: string]: any };
  themeVariant?: IReadonlyTheme;
  fields?: string;
  excludeFields?: string;
  label?: string;
  icon?: string;
  services?: {
    spHttpClient: SPHttpClient;
    pageContext: PageContext;
  };
}

interface IProperty {
  name: string;
  title: string;
  value: string;
}

interface IDocumentDetails {
  properties: IProperty[];
  system: IProperty[];
  formUrl?: string;
  fileUrl?: string;
}

interface IAdsDetailDocumentState {
  open: boolean;
  loading: boolean;
  details?: IDocumentDetails;
  error?: string;
}

/* ------------------------------------------------------------------ */
/* Outils                                                              */
/* ------------------------------------------------------------------ */

/** Lecture insensible à la casse d'une propriété du résultat de recherche. */
function getItemValue(item: { [key: string]: any }, name: string): any {
  if (!item || !name) return undefined;
  if (item[name] !== undefined && item[name] !== null && item[name] !== "") return item[name];
  const lower = name.toLowerCase();
  const key = Object.keys(item).find(k => k.toLowerCase() === lower);
  return key ? item[key] : undefined;
}

function firstValue(item: { [key: string]: any }, names: string[]): string {
  for (const n of names) {
    const v = getItemValue(item, n);
    if (v !== undefined && v !== null && String(v).trim() !== "") return String(v).trim();
  }
  return "";
}

function splitList(value?: string): string[] {
  return (value || "").split(/[,;]/).map(s => s.trim()).filter(Boolean);
}

/**
 * FieldValuesAsText encode les « _ » des noms internes en « _x005f_ »
 * (ex. Nature_x0020_doc → Nature_x005f_x0020_x005f_doc). On essaie les variantes connues.
 */
function readAsText(values: { [key: string]: any }, field: { InternalName: string; EntityPropertyName?: string }): string | undefined {
  const n = field.InternalName;
  const candidates = [
    n,
    n.replace(/_/g, "_x005f_"),
    field.EntityPropertyName,
    `OData_${n}`,
    `OData_${n.replace(/_/g, "_x005f_")}`
  ].filter(Boolean) as string[];
  for (const c of candidates) {
    if (values[c] !== undefined) return values[c];
  }
  const lowers = candidates.map(c => c.toLowerCase());
  const key = Object.keys(values).find(k => lowers.indexOf(k.toLowerCase()) > -1);
  return key ? values[key] : undefined;
}

function readJson(res: SPHttpClientResponse, what: string): Promise<any> {
  if (!res.ok) {
    return res.text().then(txt => {
      let msg = `${res.status} ${res.statusText}`;
      try { const j = JSON.parse(txt); if (j && j["odata.error"] && j["odata.error"].message) msg = j["odata.error"].message.value; } catch { /* ignore */ }
      throw new Error(`${what} : ${msg}`);
    });
  }
  return res.json();
}

/** Colonnes système affichées en bas du panneau (si présentes). */
const SYSTEM_FIELDS: Array<{ InternalName: string; title: string }> = [
  { InternalName: "ContentType", title: "Type de contenu" },
  { InternalName: "_UIVersionString", title: "Version" },
  { InternalName: "Modified", title: "Modifié le" },
  { InternalName: "Editor", title: "Modifié par" },
  { InternalName: "Created", title: "Créé le" },
  { InternalName: "Author", title: "Créé par" },
  { InternalName: "File_x0020_Size", title: "Taille (octets)" }
];

/** Colonnes jamais utiles dans le panneau (techniques ou déjà en titre). */
const ALWAYS_EXCLUDED = ["ContentType", "ContentTypeId", "_ModerationComments", "_ModerationStatus", "SelectFilename", "Attachments"];

/* ------------------------------------------------------------------ */
/* Chargement (REST SharePoint, mis en cache par document)             */
/* ------------------------------------------------------------------ */

const detailsCache: Map<string, Promise<IDocumentDetails>> = new Map();

function loadDetails(props: IAdsDetailDocumentProps): Promise<IDocumentDetails> {
  const { item, services } = props;
  const { spHttpClient, pageContext } = services;

  const webUrl = (firstValue(item, ["SPWebUrl", "SPWebURL", "SPSiteURL"]) || pageContext.web.absoluteUrl).replace(/\/$/, "");
  const listId = firstValue(item, ["NormListID", "ListID", "ListId"]).replace(/[{}]/g, "");
  const itemId = firstValue(item, ["ListItemID", "ListItemId"]);
  const fileUrl = firstValue(item, ["Path", "OriginalPath", "DefaultEncodingURL"]);

  // Accès à l'élément : par liste + ID si disponibles, sinon par l'URL du fichier
  let base: string;
  let alias = "";
  if (listId && /^\d+$/.test(itemId)) {
    base = `${webUrl}/_api/web/lists(guid'${listId}')/items(${itemId})`;
  } else if (fileUrl) {
    base = `${webUrl}/_api/web/GetFileByUrl(@u)/ListItemAllFields`;
    alias = `@u='${encodeURIComponent(fileUrl.replace(/'/g, "''"))}'`;
  } else {
    return Promise.reject(new Error("propriétés ListItemID / NormListID / Path absentes du résultat de recherche"));
  }

  const cacheKey = `${base}|${alias}`;
  if (detailsCache.has(cacheKey)) return detailsCache.get(cacheKey);

  const url = (path: string, query: string): string => {
    const q = [query, alias].filter(Boolean).join("&");
    return `${base}${path}${q ? "?" + q : ""}`;
  };
  const headers = { "Accept": "application/json;odata=nometadata" };

  const valuesPromise = spHttpClient
    .get(url("/FieldValuesAsText", ""), SPHttpClient.configurations.v1, { headers })
    .then(res => readJson(res, "Propriétés du document inaccessibles"));

  const fieldsPromise = spHttpClient
    .get(url("/ContentType/Fields", "$select=InternalName,EntityPropertyName,Title,Hidden,TypeAsString"), SPHttpClient.configurations.v1, { headers })
    .then(res => readJson(res, "Colonnes du type de contenu inaccessibles"));

  const idsPromise = spHttpClient
    .get(url("", "$select=Id,ParentList/Id,ParentList/ParentWebUrl&$expand=ParentList"), SPHttpClient.configurations.v1, { headers })
    .then(res => readJson(res, "Élément introuvable"))
    .catch(() => null); // facultatif : sert seulement au lien vers le formulaire

  const promise = Promise.all([valuesPromise, fieldsPromise, idsPromise]).then(([values, fieldsData, ids]) => {
    const ctFields: any[] = (fieldsData && fieldsData.value) || [];
    const wanted = splitList(props.fields);
    const excluded = splitList(props.excludeFields).concat(ALWAYS_EXCLUDED).map(s => s.toLowerCase());

    let fields: any[];
    if (wanted.length > 0) {
      // Ordre imposé par data-fields ; une colonne hors type de contenu reste lisible si l'élément la porte
      fields = wanted.map(name => {
        const f = ctFields.find(c => c.InternalName.toLowerCase() === name.toLowerCase());
        return f || { InternalName: name, Title: name };
      });
    } else {
      fields = ctFields.filter(f => !f.Hidden);
    }

    const properties: IProperty[] = fields
      .filter(f => excluded.indexOf(f.InternalName.toLowerCase()) === -1)
      .map(f => ({ name: f.InternalName, title: f.Title, value: (readAsText(values, f) || "").trim() }));

    const shown = properties.map(p => p.name.toLowerCase());
    const system: IProperty[] = SYSTEM_FIELDS
      .filter(f => shown.indexOf(f.InternalName.toLowerCase()) === -1)
      .map(f => ({ name: f.InternalName, title: f.title, value: (readAsText(values, f) || "").trim() }))
      .filter(p => p.value !== "");

    const formListId = (ids && ids.ParentList && ids.ParentList.Id) || listId;
    const formItemId = (ids && ids.Id !== undefined && String(ids.Id)) || itemId;
    const formWebUrl = ((ids && ids.ParentList && ids.ParentList.ParentWebUrl)
      ? `${new URL(webUrl).origin}${ids.ParentList.ParentWebUrl}`
      : webUrl).replace(/\/$/, "");
    const formUrl = formListId && formItemId
      ? `${formWebUrl}/_layouts/15/listform.aspx?PageType=4&ListId=${encodeURIComponent(`{${formListId}}`)}&ID=${formItemId}`
      : undefined;

    return { properties, system, formUrl, fileUrl: fileUrl || undefined };
  }).catch((err: any) => {
    detailsCache.delete(cacheKey); // on retentera au prochain clic
    throw err;
  });

  detailsCache.set(cacheKey, promise);
  return promise;
}

/* ------------------------------------------------------------------ */
/* Composant React                                                     */
/* ------------------------------------------------------------------ */

export class AdsDetailDocumentComponent extends React.Component<IAdsDetailDocumentProps, IAdsDetailDocumentState> {

  private _mounted = false;
  private _host = React.createRef<HTMLSpanElement>();

  constructor(props: IAdsDetailDocumentProps) {
    super(props);
    this.state = { open: false, loading: false };
  }

  public componentDidMount(): void { this._mounted = true; }
  public componentWillUnmount(): void { this._mounted = false; }

  public render(): React.ReactElement {
    const { open, loading, details, error } = this.state;
    const theme = (this.props.themeVariant as ITheme) || getTheme();
    const label = (this.props.label || "").trim() || "Détail";
    const title = firstValue(this.props.item, ["Title", "Filename"]) || "Document";

    return (
      <span ref={this._host} className="ads-detail-document" onClick={this._stop} onKeyDown={this._stopKeys}>
        <TooltipHost content={label}>
          <IconButton
            iconProps={{ iconName: (this.props.icon || "").trim() || "Info" }}
            ariaLabel={`${label} : ${title}`}
            onClick={this._onOpen}
            theme={theme}
          />
        </TooltipHost>
        <Panel
          isOpen={open}
          type={PanelType.medium}
          isLightDismiss={true}
          headerText={title}
          closeButtonAriaLabel="Fermer"
          onDismiss={this._onClose}
          theme={theme}
          onRenderFooterContent={details ? this._renderFooter : undefined}
          isFooterAtBottom={true}
        >
          {loading && <Spinner size={SpinnerSize.medium} label="Chargement des propriétés…" styles={{ root: { marginTop: 24 } }} />}
          {error && (
            <MessageBar messageBarType={MessageBarType.error} isMultiline={true} styles={{ root: { marginTop: 16 }, text: { userSelect: "text", wordBreak: "break-word" } }}>
              {error}
            </MessageBar>
          )}
          {details && this._renderProperties(details, theme)}
        </Panel>
      </span>
    );
  }

  private _renderProperties(details: IDocumentDetails, theme: ITheme): React.ReactElement {
    const row = (p: IProperty): React.ReactElement => (
      <div key={p.name} style={{ padding: "10px 0", borderBottom: `1px solid ${theme.palette.neutralLighter}` }}>
        <div style={{ fontSize: 12, fontWeight: 600, color: theme.palette.neutralSecondary, marginBottom: 2 }}>{p.title}</div>
        <div style={{ fontSize: 14, color: p.value ? theme.palette.neutralPrimary : theme.palette.neutralTertiary, whiteSpace: "pre-wrap", wordBreak: "break-word", userSelect: "text" }}>
          {p.value || "—"}
        </div>
      </div>
    );
    return (
      <div style={{ marginTop: 8 }}>
        {details.properties.map(row)}
        {details.system.length > 0 && (
          <>
            <div style={{ marginTop: 20, fontSize: 12, fontWeight: 700, letterSpacing: ".06em", textTransform: "uppercase", color: theme.palette.themePrimary }}>
              Informations
            </div>
            {details.system.map(row)}
          </>
        )}
      </div>
    );
  }

  private _renderFooter = (): JSX.Element => {
    const { details } = this.state;
    const theme = (this.props.themeVariant as ITheme) || getTheme();
    return (
      <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
        {details.fileUrl && (
          <ActionButton iconProps={{ iconName: "OpenInNewWindow" }} href={details.fileUrl} target="_blank" rel="noopener" data-interception="off" theme={theme}>
            Ouvrir le document
          </ActionButton>
        )}
        {details.formUrl && (
          <ActionButton iconProps={{ iconName: "Edit" }} href={details.formUrl} target="_blank" rel="noopener" data-interception="off" theme={theme}>
            Formulaire SharePoint
          </ActionButton>
        )}
        {!details.fileUrl && !details.formUrl && <Link onClick={this._onClose}>Fermer</Link>}
      </div>
    );
  }

  /**
   * Le bouton est placé dans un <a> (carte) : on bloque la navigation pour les clics sur l'icône.
   * Les clics dans le panneau (rendu dans un calque hors de la carte) remontent aussi ici via React :
   * on les laisse passer pour que les liens du panneau fonctionnent.
   */
  private _stop = (e: React.MouseEvent<any>): void => {
    if (this._host.current && this._host.current.contains(e.target as Node)) {
      e.preventDefault();
      e.stopPropagation();
    }
  }

  private _stopKeys = (e: React.KeyboardEvent<any>): void => {
    // Entrée/Espace sur l'icône ne doivent pas activer le lien de la carte
    if (this._host.current && this._host.current.contains(e.target as Node) && (e.key === "Enter" || e.key === " ")) {
      e.stopPropagation();
    }
  }

  private _onOpen = (e?: any): void => {
    if (e) { e.preventDefault(); e.stopPropagation(); }
    this.setState({ open: true });
    if (this.state.details || this.state.loading) return;

    if (!this.props.item || !this.props.services) {
      this.setState({ error: "Données du résultat indisponibles (attribut data-item absent du template ?)" });
      return;
    }

    this.setState({ loading: true, error: undefined });
    loadDetails(this.props)
      .then(details => { if (this._mounted) this.setState({ loading: false, details }); })
      .catch((err: any) => {
        Log.error(LOG_SOURCE, err);
        if (this._mounted) this.setState({ loading: false, error: `Impossible d'afficher les propriétés : ${err && err.message ? err.message : err}` });
      });
  }

  private _onClose = (): void => {
    this.setState({ open: false });
  }
}

/* ------------------------------------------------------------------ */
/* Web component                                                       */
/* ------------------------------------------------------------------ */

export class AdsDetailDocumentWebComponent extends BaseWebComponent {

  public constructor() {
    super();
  }

  public connectedCallback(): void {
    try {
      const props: IAdsDetailDocumentProps = this.resolveAttributes();
      props.services = {
        spHttpClient: this._serviceScope.consume(SPHttpClient.serviceKey),
        pageContext: this._serviceScope.consume(PageContext.serviceKey)
      };
      ReactDOM.render(<AdsDetailDocumentComponent {...props} />, this);
    } catch (err) {
      Log.error(LOG_SOURCE, err);
      // eslint-disable-next-line no-console
      console.error("[ADS Détail] erreur d'initialisation", err);
    }
  }

  protected onDispose(): void {
    ReactDOM.unmountComponentAtNode(this);
  }
}
