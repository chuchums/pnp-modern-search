import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseWebComponent, IDataFilterInfo, IDataFilterInternal, IDataFilterValueInfo, IDataFilterValueInternal, ExtensibilityConstants } from '@pnp/modern-search-extensibility';
import { ITheme, Text, Toggle, getTheme } from '@fluentui/react';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { TaxonomyHelper } from '../../helpers/TaxonomyHelper';

/**
 * Values considered as "true" for a Yes/No (boolean) managed property refiner.
 * SharePoint usually returns "1"/"0", but "true"/"false" may appear depending on the mapping.
 */
const TRUTHY_VALUES = ['1', 'true', 'yes', 'oui', 'vrai', '-1'];

const setImmediateProgressCursor = (): void => {
    if (!globalThis.document) {
        return;
    }

    if (globalThis.document.documentElement) {
        globalThis.document.documentElement.style.setProperty('cursor', 'progress', 'important');
    }

    if (globalThis.document.body) {
        globalThis.document.body.style.setProperty('cursor', 'progress', 'important');
    }

    const styleId = 'pnp-modern-search-busy-cursor-style';
    if (!globalThis.document.getElementById(styleId)) {
        const styleElement = globalThis.document.createElement('style');
        styleElement.id = styleId;
        styleElement.textContent = '* { cursor: progress !important; }';
        globalThis.document.head.appendChild(styleElement);
    }
};

export interface IFilterBooleanProps {

    /**
     * The filter information (name, display name, values) as passed by the Handlebars template
     */
    filter: IDataFilterInternal;

    /**
     * The Web Part instance ID from where the filter component belongs
     */
    instanceId?: string;

    /**
     * Optional label displayed next to the toggle. Defaults to the filter display name.
     */
    label?: string;

    /**
     * The current theme settings
     */
    themeVariant?: IReadonlyTheme;

    /**
     * Handler when the toggle state changes
     */
    onChanged: (filterName: string, filterValue: IDataFilterValueInfo) => void;
}

export class FilterBooleanComponent extends React.Component<IFilterBooleanProps, {}> {

    /**
     * Normalizes a raw refiner name/value ("1", "\"ǂǂ31\"", "True"...) to a lowercase plain string
     */
    public static normalize(raw: string): string {
        const value = `${raw ?? ''}`.trim().replace(/^"+|"+$/g, '');
        if (!value) {
            return '';
        }

        if (value.indexOf('ǂ') !== -1) {
            const decoded = TaxonomyHelper.decodeHexString(value);
            if (decoded) {
                return decoded.trim().toLowerCase();
            }
        }

        return value.toLowerCase();
    }

    public static isTruthy(filterValue: { name?: string; value?: string }): boolean {
        return TRUTHY_VALUES.indexOf(FilterBooleanComponent.normalize(filterValue?.name)) !== -1
            || TRUTHY_VALUES.indexOf(FilterBooleanComponent.normalize(filterValue?.value)) !== -1;
    }

    /**
     * Returns the refiner value matching "true" in the current results (if any)
     */
    public static getTrueValue(filter: IDataFilterInternal): IDataFilterValueInternal | undefined {
        return (filter?.values || []).filter(value => FilterBooleanComponent.isTruthy(value))[0];
    }

    private _getTextColor(): string {
        if (this.props.themeVariant?.isInverted) {
            return this.props.themeVariant?.semanticColors?.bodyText ?? 'inherit';
        }
        return this.props.themeVariant?.semanticColors?.inputText ?? 'inherit';
    }

    /**
     * On dark/colored sections (or without theme) the toggle follows the text color
     */
    private _useTextColorForToggle(): boolean {
        return !this.props.themeVariant?.semanticColors || !!this.props.themeVariant?.isInverted;
    }

    public render(): JSX.Element {

        const filter = this.props.filter;
        const trueValue = FilterBooleanComponent.getTrueValue(filter);
        const isChecked = !!trueValue?.selected;
        const count = trueValue?.count ?? 0;

        // Nothing to filter on when no result has the value "true" (unless already selected)
        const isDisabled = !trueValue || (!isChecked && count === 0);

        const textColor = this._getTextColor();
        const label = this.props.label?.trim() ? this.props.label : filter.displayName;
        const countLabel = filter.showCount && trueValue ? ` (${count})` : '';
        const toggleId = `pnp-filterboolean-${this.props.instanceId || ''}-${filter.filterName}`.replace(/[^a-zA-Z0-9_-]/g, '-');
        const useTextColor = this._useTextColorForToggle();

        return (
            <div
                style={{
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    gap: 12,
                    width: '100%',
                    padding: '8px 0',
                    opacity: isDisabled ? 0.6 : 1
                }}
                title={isDisabled ? undefined : `${label}${countLabel}`}
            >
                <label htmlFor={toggleId} style={{ cursor: isDisabled ? 'default' : 'pointer', minWidth: 0, flex: 1 }}>
                    <Text variant={'large'} block nowrap styles={{ root: { color: textColor } }}>
                        {label}
                        {countLabel ? <Text variant={'medium'} styles={{ root: { color: textColor, marginLeft: 4 } }}>{countLabel}</Text> : null}
                    </Text>
                </label>
                <Toggle
                    id={toggleId}
                    key={`${filter.filterName}-${isChecked}`}
                    defaultChecked={isChecked}
                    disabled={isDisabled}
                    role="switch"
                    aria-label={`${label}${countLabel}`}
                    theme={(this.props.themeVariant as ITheme) || getTheme()}
                    styles={{
                        root: { marginBottom: 0 },
                        pill: useTextColor ? {
                            borderColor: textColor,
                            selectors: {
                                ':hover': { borderColor: textColor }
                            }
                        } : undefined,
                        thumb: useTextColor && !isChecked ? { backgroundColor: textColor } : undefined
                    }}
                    onChange={(ev, checked: boolean) => {
                        if (!trueValue) {
                            return;
                        }
                        setImmediateProgressCursor();
                        this.props.onChanged(filter.filterName, {
                            name: trueValue.name,
                            value: trueValue.value,
                            selected: checked
                        });
                    }}
                />
            </div>
        );
    }
}

export class FilterBooleanWebComponent extends BaseWebComponent {

    public constructor() {
        super();
    }

    public connectedCallback() {

        const props = this.resolveAttributes();

        if (!props.filter) {
            return;
        }

        const filter = props.filter as IDataFilterInternal;

        const toggle = <FilterBooleanComponent {...props} filter={filter} onChanged={(filterName: string, filterValue: IDataFilterValueInfo) => {

            // Make sure the other value ("0") is never left selected
            const otherValues: IDataFilterValueInfo[] = (filter.values || [])
                .filter(value => value.value !== filterValue.value && value.selected)
                .map(value => ({ name: value.name, value: value.value, selected: false }));

            const detail: IDataFilterInfo = {
                filterName: filterName,
                filterValues: [filterValue, ...otherValues],
                instanceId: props.instanceId,
                // Apply immediately, no "Apply" button for a toggle
                forceUpdate: true
            };

            this.dispatchEvent(new CustomEvent(ExtensibilityConstants.EVENT_FILTER_UPDATED, {
                detail,
                bubbles: true,
                cancelable: true
            }));
        }}
        />;

        ReactDOM.render(toggle, this);
    }

    protected onDispose(): void {
        ReactDOM.unmountComponentAtNode(this);
    }
}
