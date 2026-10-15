import * as React from "react";
import * as ReactDOM from "react-dom";
import { BaseWebComponent } from "@pnp/modern-search-extensibility";
import {
  ActionButton, DefaultButton, Dialog, DialogFooter, DialogType, IIconProps, ITheme,
  MessageBar, MessageBarType, PrimaryButton, Spinner, SpinnerSize, getTheme
} from "@fluentui/react";
import { AadTokenProviderFactory, HttpClient, HttpClientResponse, SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { PageContext } from "@microsoft/sp-page-context";
import { Log } from "@microsoft/sp-core-library";
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { IAdsEnvoiRedactionSettings } from "../models/common/IAdsEnvoiRedactionSettings";

/**
 * ADS GED – bouton « Envoi en rédaction » pour les résultats PnP Modern Search.
 *
 * Toute la configuration vient de l'onglet « Envoi en rédaction » de la webpart Résultats
 * (propriété adsEnvoiRedaction). Dans le template Handlebars, il suffit de :
 *
 *   <ads-envoi-redaction
 *       data-config="{{JSONstringify @root.properties.adsEnvoiRedaction}}"
 *       data-item="{{JSONstringify item}}"
 *       data-theme-variant="{{JSONstringify @root.theme}}">
 *   </ads-envoi-redaction>
 *
 * Le bouton s'affiche si l'utilisateur courant est dans la colonne « Rédacteurs » d'une ligne
 * de la liste de paramétrage dont les colonnes mappées correspondent aux propriétés gérées du document.
 */

const LOG_SOURCE = "AdsEnvoiRedaction";
const FLOW_RESOURCE = "https://service.flow.microsoft.com/";
/** Passerelle des connecteurs Power Platform (azure-apim.net) */
const APIHUB_RESOURCE = "https://apihub.azure.com";

export interface IAdsEnvoiRedactionProps {
  config?: IAdsEnvoiRedactionSettings;
  item?: { [key: string]: any };
  themeVariant?: IReadonlyTheme;
  services?: {
    spHttpClient: SPHttpClient;
    httpClient: HttpClient;
    aadTokenProviderFactory: AadTokenProviderFactory;
    pageContext: PageContext;
  };
}

interface IAdsEnvoiRedactionState {
  checking: boolean;
  allowed: boolean;
  confirmOpen: boolean;
  sending: boolean;
  sent: boolean;
  error?: string;
  /** Mode diagnostic (?adsDebug=1) : explication de la décision */
  debugLines?: string[];
}

/** Une ligne de la liste de paramétrage où l'utilisateur est rédacteur : jetons par colonne mappée. */
type IConfigRule = Array<Set<string>>;

/* ------------------------------------------------------------------ */
/* Normalisation des valeurs (texte, choix, lookup, taxonomie, search) */
/* ------------------------------------------------------------------ */

const GUID_RE = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i;

export function toTokens(value: any): string[] {
  if (value === null || value === undefined || value === "") return [];
  if (Array.isArray(value)) return value.reduce((acc: string[], v) => acc.concat(toTokens(v)), []);

  if (typeof value === "object") {
    const out: string[] = [];
    if (value.TermID) out.push(String(value.TermID));            // taxonomie (RenderListDataAsStream)
    if (value.TermGuid) out.push(String(value.TermGuid));
    if (value.Label && isNaN(Number(value.Label))) out.push(String(value.Label));
    if (value.lookupValue) out.push(String(value.lookupValue));  // lookup
    if (value.Title) out.push(String(value.Title));
    if (value.title) out.push(String(value.title));              // personne
    if (value.email) out.push(String(value.email));
    return out.map(normalize).filter(Boolean);
  }

  // Chaînes : multi-valeurs séparées par ';' ou retour ligne, taxonomie search "L0|#0<guid>|Libellé"
  const out: string[] = [];
  String(value).split(/;|\n/).forEach(rawPart => {
    const part = rawPart.trim();
    if (!part) return;
    if (part.indexOf("|#") > -1) {
      const prefix = part.substring(0, part.indexOf("|#")).toUpperCase();
      if (prefix === "GTSET" || prefix === "GPP") return; // jeu de termes / termes parents : ignorés
      const guid = part.match(GUID_RE);
      if (guid) out.push(guid[0]);
      const pipe = part.lastIndexOf("|");
      if (prefix === "L0" && pipe > part.indexOf("|#") + 1) out.push(part.substring(pipe + 1));
      return;
    }
    out.push(part);
  });
  return out.map(normalize).filter(Boolean);
}

/** Minuscules, sans accents ni espaces superflus : « Contrôle  de gestion » = « controle de gestion ». */
function normalize(t: string): string {
  return String(t).normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/\s+/g, " ").trim().toLowerCase();
}

/** Mode diagnostic : ajouter ?adsDebug=1 à l'URL de la page. */
function isDebug(): boolean {
  try { return /[?&]adsDebug=1/i.test(window.location.search); } catch { return false; }
}

/** Lecture insensible à la casse d'une propriété du résultat de recherche. */
function getItemValue(item: { [key: string]: any }, name: string): any {
  if (!item || !name) return undefined;
  if (item[name] !== undefined) return item[name];
  const lower = name.toLowerCase();
  const key = Object.keys(item).find(k => k.toLowerCase() === lower);
  return key ? item[key] : undefined;
}

/* ------------------------------------------------------------------ */
/* Chargement de la liste de paramétrage (1 appel par page, mis en cache) */
/* ------------------------------------------------------------------ */

interface IRulesResult {
  rules: IConfigRule[];
  diag: {
    user?: any;
    totalRows?: number;
    rawFirstRow?: any;
    error?: string;
  };
}

const rulesCache: Map<string, Promise<IRulesResult>> = new Map();

function readJson(res: SPHttpClientResponse, what: string): Promise<any> {
  if (!res.ok) throw new Error(`${what} (${res.status})`);
  return res.json();
}

function loadRulesForCurrentUser(config: IAdsEnvoiRedactionSettings, services: IAdsEnvoiRedactionProps["services"]): Promise<IRulesResult> {
  const { spHttpClient, pageContext } = services;
  const siteUrl = (config.configSiteUrl || pageContext.web.absoluteUrl).replace(/\/$/, "");
  const mappings = (config.mappings || []).filter(m => m && m.listField && m.managedProperty);
  const fields = mappings.map(m => m.listField).concat([config.redacteursField]);
  const cacheKey = `${siteUrl}|${config.configListId}|${fields.join(",")}`;

  if (rulesCache.has(cacheKey)) return rulesCache.get(cacheKey);

  const jsonHeaders = { "Accept": "application/json;odata=nometadata", "Content-Type": "application/json;odata=nometadata" };
  const viewXml =
    `<View Scope="RecursiveAll"><ViewFields>${fields.map(f => `<FieldRef Name="${f}"/>`).join("")}</ViewFields>` +
    `<RowLimit Paged="TRUE">5000</RowLimit></View>`;

  // Utilisateur courant tel que connu du site de la liste : son Id est comparé à l'id renvoyé par la colonne Personne
  const currentUserPromise = spHttpClient
    .get(`${siteUrl}/_api/web/currentuser?$select=Id,Email,UserPrincipalName,LoginName,Title`, SPHttpClient.configurations.v1, { headers: jsonHeaders })
    .then(res => readJson(res, "Utilisateur courant introuvable sur le site de paramétrage"));

  const rowsPromise = spHttpClient
    .post(
      `${siteUrl}/_api/web/lists(guid'${config.configListId}')/RenderListDataAsStream`,
      SPHttpClient.configurations.v1,
      { headers: jsonHeaders, body: JSON.stringify({ parameters: { RenderOptions: 2, ViewXml: viewXml, DatesInUtc: true } }) }
    )
    .then(res => readJson(res, "Liste de paramétrage inaccessible"));

  const promise: Promise<IRulesResult> = Promise.all([currentUserPromise, rowsPromise])
    .then(([user, data]) => {
      const rows: any[] = (data && data.Row) || [];
      const userId = user && user.Id !== undefined ? String(user.Id) : null;
      const userKeys = [user && user.Email, user && user.UserPrincipalName, pageContext.user.email, pageContext.user.loginName]
        .filter(Boolean)
        .map((s: string) => s.toLowerCase());

      const rules = rows
        .filter(row => {
          const value = row[config.redacteursField];
          const redacteurs: any[] = Array.isArray(value) ? value : (value ? [value] : []);
          return redacteurs.some(u => {
            if (userId && u.id !== undefined && String(u.id) === userId) return true;
            const candidates = [u.email, u.EMail, u.sip, u.loginName].filter(Boolean).map((s: string) => s.toLowerCase());
            return candidates.some(c => userKeys.some(k => c === k || c.endsWith(`|${k}`)));
          });
        })
        .map(row => mappings.map(m => new Set(toTokens(row[m.listField]))));

      return { rules, diag: { user, totalRows: rows.length, rawFirstRow: rows[0] } };
    })
    .catch((err: any) => {
      Log.error(LOG_SOURCE, err);
      rulesCache.delete(cacheKey); // on retentera au prochain rendu
      return { rules: [] as IConfigRule[], diag: { error: `${err && err.message ? err.message : err}` } };
    });

  rulesCache.set(cacheKey, promise);
  return promise;
}

/* ------------------------------------------------------------------ */
/* Composant React                                                     */
/* ------------------------------------------------------------------ */

export class AdsEnvoiRedactionComponent extends React.Component<IAdsEnvoiRedactionProps, IAdsEnvoiRedactionState> {

  private _mounted = false;

  constructor(props: IAdsEnvoiRedactionProps) {
    super(props);
    this.state = { checking: true, allowed: false, confirmOpen: false, sending: false, sent: false };
    this._onClick = this._onClick.bind(this);
    this._onConfirm = this._onConfirm.bind(this);
    this._onCancel = this._onCancel.bind(this);
  }

  public componentDidMount(): void {
    this._mounted = true;
    this._checkPermission();
  }

  public componentWillUnmount(): void {
    this._mounted = false;
  }

  public render(): React.ReactElement {
    const { checking, allowed, confirmOpen, sending, sent, error, debugLines } = this.state;

    const debugPanel = debugLines && debugLines.length > 0
      ? (
        <pre onClick={this._stop} style={{ whiteSpace: "pre-wrap", fontSize: 11, lineHeight: 1.4, margin: "6px 0 0", padding: 8, background: "#FFF8E1", border: "1px solid #E6C200", color: "#333", cursor: "text", userSelect: "text" }}>
          {debugLines.join("\n")}
        </pre>
      ) : null;

    if (checking) return null;
    if (!allowed) return debugPanel;

    const config = this.props.config;
    const label = (config.buttonLabel || "").trim() || "Envoi en rédaction";
    const theme = (this.props.themeVariant as ITheme) || getTheme();
    const icon: IIconProps = { iconName: sent ? "CheckMark" : "Send" };
    const fileName = this._fileName();

    return (
      <span className="ads-envoi-redaction" onClick={this._stop}>
        <ActionButton
          text={label}
          iconProps={icon}
          disabled={sending || sent}
          onClick={this._onClick}
          theme={theme}
        />
        {sending && <Spinner size={SpinnerSize.xSmall} styles={{ root: { display: "inline-block", marginLeft: 6 } }} />}
        {error && (
          <MessageBar messageBarType={MessageBarType.error} isMultiline={true} styles={{ text: { userSelect: "text", wordBreak: "break-word" } }} onDismiss={() => this.setState({ error: undefined })}>
            {error}
          </MessageBar>
        )}
        <Dialog
          hidden={!confirmOpen}
          onDismiss={this._onCancel}
          dialogContentProps={{
            type: DialogType.normal,
            title: label,
            subText: (config.confirmText || "").trim()
              ? config.confirmText.replace("{fileName}", fileName || "ce document")
              : `Confirmer « ${label} » pour « ${fileName || "ce document"} » ?`
          }}
          modalProps={{ isBlocking: false }}
        >
          <DialogFooter>
            <PrimaryButton onClick={this._onConfirm} text="Confirmer" />
            <DefaultButton onClick={this._onCancel} text="Annuler" />
          </DialogFooter>
        </Dialog>
        {debugPanel}
      </span>
    );
  }

  /** Le bouton peut être placé dans un <a> (carte) : on empêche la navigation. */
  private _stop = (e: React.MouseEvent<any>): void => {
    e.preventDefault();
    e.stopPropagation();
  }

  private _fileName(): string {
    const item = this.props.item;
    return getItemValue(item, "Filename") || getItemValue(item, "Title") || "";
  }

  private _checkPermission(): void {
    const { config, item, services } = this.props;
    const debug = isDebug();
    const mappings = ((config && config.mappings) || []).filter(m => m && m.listField && m.managedProperty);

    const missing: string[] = [];
    if (!config) missing.push("data-config vide (attribut absent du template ?)");
    else {
      if (!config.enabled) missing.push("bouton désactivé dans la webpart");
      if (!config.configListId) missing.push("liste de paramétrage non choisie");
      if (!config.redacteursField) missing.push("colonne Rédacteurs non choisie");
      if (mappings.length === 0) missing.push("aucune correspondance colonne ↔ propriété gérée");
    }
    if (!item) missing.push("data-item vide");
    if (!services) missing.push("services SPFx indisponibles");

    if (missing.length > 0) {
      this.setState({ checking: false, allowed: false, debugLines: debug ? ["[ADS] Bouton masqué : " + missing.join(" ; ")] : undefined });
      return;
    }

    const docRaw = mappings.map(m => getItemValue(item, m.managedProperty));
    const docTokens: string[][] = docRaw.map(v => toTokens(v));
    const matchAll = config.matchMode === "all";

    loadRulesForCurrentUser(config, services).then(({ rules, diag }) => {
      const allowed = rules.some(rule => {
        // Colonnes renseignées sur cette ligne de paramétrage (colonne vide = pas de contrainte)
        const checks = rule
          .map((rowTokens, i) => ({ rowTokens, docTokens: docTokens[i] }))
          .filter(c => c.rowTokens.size > 0);
        if (checks.length === 0) return false;
        const ok = (c: { rowTokens: Set<string>, docTokens: string[] }): boolean => c.docTokens.some(t => c.rowTokens.has(t));
        return matchAll ? checks.every(ok) : checks.some(ok);
      });

      let debugLines: string[];
      if (debug) {
        const u = diag.user || {};
        debugLines = [
          `[ADS] ${allowed ? "AFFICHÉ" : "MASQUÉ"} – mode ${matchAll ? "toutes" : "au moins une"}`,
          diag.error ? `Erreur : ${diag.error}` : `Utilisateur (site de paramétrage) : Id=${u.Id} ; Email=${u.Email} ; UPN=${u.UserPrincipalName}`,
          `Lignes de paramétrage : ${diag.totalRows ?? "?"} au total, ${rules.length} où vous êtes rédacteur`,
          ...mappings.map((m, i) =>
            `• ${m.listField} ↔ ${m.managedProperty} | document : ${JSON.stringify(docRaw[i])} → [${docTokens[i].join(" | ")}]` +
            ` | vos lignes : ${rules.map(r => `[${Array.from(r[i]).join(" | ")}]`).join(" ") || "aucune"}`
          ),
          rules.length === 0 && diag.rawFirstRow
            ? `1re ligne brute : ${JSON.stringify(diag.rawFirstRow[config.redacteursField])}`
            : ""
        ].filter(Boolean);
        // eslint-disable-next-line no-console
        console.log("[ADS Envoi en rédaction]", { item, config, diag, docTokens, rules: rules.map(r => r.map(s => Array.from(s))), allowed });
      }

      if (this._mounted) this.setState({ checking: false, allowed, debugLines });
    });
  }

  private _onClick(e?: any): void {
    if (e) { e.preventDefault(); e.stopPropagation(); }
    this.setState({ confirmOpen: true, error: undefined });
  }

  private _onCancel(): void {
    this.setState({ confirmOpen: false });
  }

  private _onConfirm(): void {
    this.setState({ confirmOpen: false, sending: true, error: undefined });

    this._runFlow()
      .then(() => { if (this._mounted) this.setState({ sending: false, sent: true }); })
      .catch((err: any) => {
        Log.error(LOG_SOURCE, err);
        if (this._mounted) this.setState({ sending: false, error: `Échec du lancement du flux : ${err && err.message ? err.message : err}` });
      });
  }

  /** Même charge utile que celle envoyée par SharePoint pour un déclencheur « Pour un fichier sélectionné ». */
  private _buildPayload(): any {
    const item = this.props.item;
    const config = this.props.config;
    const idValue = getItemValue(item, "ListItemID");
    const id = idValue !== undefined && idValue !== null ? String(idValue) : "";
    const fileName = this._fileName();
    const fileUrl = getItemValue(item, "Path") || getItemValue(item, "OriginalPath") || "";

    // Variables disponibles dans le JSON complémentaire : {ID} {itemUrl} {fileName} {userEmail} {userName}
    const vars: { [k: string]: string } = {
      ID: id,
      itemUrl: fileUrl,
      fileName,
      userEmail: this.props.services.pageContext.user.email || "",
      userName: this.props.services.pageContext.user.displayName || ""
    };

    let extra: any = {};
    const extraRaw = (config.flowExtraBody || "").trim();
    if (extraRaw) {
      const replaced = extraRaw.replace(/\{(ID|itemUrl|fileName|userEmail|userName)\}/g, (_m, k) => JSON.stringify(vars[k]).slice(1, -1));
      try { extra = JSON.parse(replaced); } catch { throw new Error("le JSON complémentaire du flux (webpart) est invalide"); }
    }

    // Champs envoyés par SharePoint pour « Pour un fichier sélectionné », complétés par le JSON de la webpart
    // Le schéma du déclencheur type ID en "integer" : on l'envoie en nombre
    const numericId = /^\d+$/.test(id) ? parseInt(id, 10) : id;
    // Identique à l'appel de SharePoint : {"rows":[{"entity":{"ID":3468,"itemUrl":"…","fileName":"…","FileId":"3468"}}]}
    const entity = {
      ID: numericId,
      itemUrl: fileUrl,
      fileName,
      FileId: id,
      ...(extra.entity || {})
    };
    const extraRoot = { ...extra };
    delete extraRoot.entity;

    return {
      rows: [{ entity }],
      ...extraRoot
    };
  }

  private _runFlow(): Promise<void> {
    const { config } = this.props;
    const { httpClient, aadTokenProviderFactory, pageContext } = this.props.services;
    let body: string;
    try { body = JSON.stringify(this._buildPayload()); } catch (e) { return Promise.reject(e); }
    // eslint-disable-next-line no-console
    if (isDebug()) console.log("[ADS Envoi en rédaction] corps envoyé au flux", JSON.parse(body));

    // Option B : déclencheur HTTP (URL signée)
    if ((config.flowUrl || "").trim()) {
      return httpClient.post(config.flowUrl.trim(), HttpClient.configurations.v1, {
        headers: { "Content-Type": "application/json" },
        body
      }).then(this._ensureOk);
    }

    // Option A : API Power Automate, au nom de l'utilisateur
    let flowId = (config.flowId || "").trim();
    let envId = (config.environmentId || "").trim();
    if (!flowId) return Promise.reject(new Error("aucun flux paramétré dans la webpart"));
    const parts = flowId.split("/");
    if (parts.length === 3 && parts[0].toLowerCase() === "v1") { envId = envId || parts[1]; flowId = parts[2]; }
    if (!envId) envId = `Default-${pageContext.aadInfo.tenantId.toString()}`;
    const host = ((config.flowApiHost || "").trim() || "https://api.flow.microsoft.com").replace(/\/$/, "");

    const triggerBase = `${host}/providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(envId)}/flows/${encodeURIComponent(flowId)}/triggers/manual`;
    const debug = isDebug();
    const log = (...args: any[]): void => {
      // eslint-disable-next-line no-console
      if (debug) console.log("[ADS Envoi en rédaction]", ...args);
    };

    return aadTokenProviderFactory.getTokenProvider()
      .then(async provider => {
        const token = await provider.getToken(FLOW_RESOURCE);

        // 0) Voie privilégiée : point d'entrée « logicflows » de la passerelle des connecteurs
        //    (https://<région>.azure-apim.net/apim/logicflows/<id>/triggers/manual/run), celui
        //    qu'utilise Power Automate lui-même. Il transmet le corps tel quel au déclencheur.
        let triggerUri = (config.flowTriggerUri || "").trim();
        if (!triggerUri) {
          try {
            const flowRes = await fetch(`${host}/providers/Microsoft.ProcessSimple/environments/${encodeURIComponent(envId)}/flows/${encodeURIComponent(flowId)}?api-version=2016-11-01`, {
              headers: { "Accept": "application/json", "Authorization": `Bearer ${token}` }
            });
            if (flowRes.ok) {
              const flow = await flowRes.json();
              triggerUri = (flow && flow.properties && flow.properties.flowTriggerUri) || "";
            }
            log("URL logicflows découverte", triggerUri || "(aucune)");
          } catch (e) { log("découverte de l'URL logicflows impossible", e); }
        }

        let apimError = "";
        if (triggerUri) {
          if (!/api-version=/i.test(triggerUri)) triggerUri += (triggerUri.indexOf("?") > -1 ? "&" : "?") + "api-version=2016-11-01";
          // Jeton pour la passerelle des connecteurs, à défaut celui du service Flow
          let apimToken = token;
          try { apimToken = await provider.getToken(APIHUB_RESOURCE); } catch (e) { log("jeton apihub indisponible, essai avec le jeton Flow", e); }
          log("appel logicflows", triggerUri);
          const apimRes = await fetch(triggerUri, {
            method: "POST",
            headers: { "Content-Type": "application/json", "Accept": "application/json, text/plain, */*", "Authorization": `Bearer ${apimToken}` },
            body
          });
          if (apimRes.ok) return;
          apimError = await this._readError(apimRes);
          log("logicflows refusé", apimRes.status, apimError);
        }

        // 1) Même appel que SharePoint (fetch natif : pas d'en-têtes ajoutés par SPFx)
        const runUrl = `${triggerBase}/run?api-version=2016-11-01`;
        log("appel run", runUrl);
        const runRes = await fetch(runUrl, {
          method: "POST",
          headers: { "Content-Type": "application/json; charset=utf-8", "Accept": "application/json", "Authorization": `Bearer ${token}` },
          body
        });
        if (runRes.ok) return;
        const runError = await this._readError(runRes);
        log("run refusé", runRes.status, runError);
        const prefix = apimError ? `logicflows : ${apimError} | run : ` : "";
        if (runRes.status !== 400) throw new Error(prefix + runError);

        // 2) Repli : URL signée du déclencheur (évite le relais Azure Resource Manager)
        const cbRes = await fetch(`${triggerBase}/listCallbackUrl?api-version=2016-11-01`, {
          method: "POST",
          headers: { "Accept": "application/json", "Authorization": `Bearer ${token}` }
        });
        if (!cbRes.ok) {
          const cbError = await this._readError(cbRes);
          log("listCallbackUrl refusé", cbRes.status, cbError);
          throw new Error(`${prefix}${runError} (repli impossible : ${cbError})`);
        }
        const cb = await cbRes.json();
        const callbackUrl: string = (cb && cb.response && cb.response.value) || (cb && cb.value) || "";
        if (!callbackUrl) throw new Error(`${prefix}${runError} (repli impossible : URL de rappel absente)`);
        log("repli via URL de rappel du déclencheur");

        const cbRun = await fetch(callbackUrl, {
          method: "POST",
          headers: { "Content-Type": "application/json; charset=utf-8" },
          body
        });
        if (!cbRun.ok) throw new Error(await this._readError(cbRun));
      });
  }

  private async _readError(res: Response | HttpClientResponse): Promise<string> {
    let msg = `${res.status} ${res.statusText}`;
    try {
      const txt = await res.text();
      try { const j = JSON.parse(txt); if (j && j.error && j.error.message) msg = j.error.message; else if (txt) msg = txt; } catch { if (txt) msg = txt; }
    } catch { /* ignore */ }
    return msg;
  }

  private _ensureOk = (res: HttpClientResponse): Promise<void> => {
    if (res.ok) return Promise.resolve();
    return this._readError(res).then(msg => { throw new Error(msg); });
  }
}

/* ------------------------------------------------------------------ */
/* Web component                                                       */
/* ------------------------------------------------------------------ */

export class AdsEnvoiRedactionWebComponent extends BaseWebComponent {

  public constructor() {
    super();
  }

  public connectedCallback(): void {
    const debug = isDebug();
    try {
      const props: IAdsEnvoiRedactionProps = this.resolveAttributes();
      // eslint-disable-next-line no-console
      if (debug) console.log("[ADS Envoi en rédaction] composant initialisé", { config: props.config, hasItem: !!props.item });

      props.services = {
        spHttpClient: this._serviceScope.consume(SPHttpClient.serviceKey),
        httpClient: this._serviceScope.consume(HttpClient.serviceKey),
        aadTokenProviderFactory: this._serviceScope.consume(AadTokenProviderFactory.serviceKey),
        pageContext: this._serviceScope.consume(PageContext.serviceKey)
      };

      ReactDOM.render(<AdsEnvoiRedactionComponent {...props} />, this);
    } catch (err) {
      Log.error(LOG_SOURCE, err);
      // eslint-disable-next-line no-console
      console.error("[ADS Envoi en rédaction] erreur d'initialisation", err);
      if (debug) this.textContent = `[ADS] Erreur d'initialisation : ${err && err.message ? err.message : err}`;
    }
  }

  protected onDispose(): void {
    ReactDOM.unmountComponentAtNode(this);
  }
}
