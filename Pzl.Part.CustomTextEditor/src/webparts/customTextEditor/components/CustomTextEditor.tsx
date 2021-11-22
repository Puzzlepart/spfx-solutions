import * as React from 'react';
import styles from './CustomTextEditor.module.scss';
import { ICustomTextEditorProps } from './ICustomTextEditorProps';
import { ICustomTextEditorState } from './ICustomTextEditorState';
import ReactHtmlParser from 'react-html-parser';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import * as strings from 'CustomTextEditorWebPartStrings';
import { TextBoxStyle } from './TextBoxStyle';
import { DisplayMode } from '@microsoft/sp-core-library';

/**
 * TinyMCE Class that contains a TinyMCE Editor instance.
 * Takes the HTML stored in the WebPart,
 * the display mode of the webpart and either dislays
 * the HTML as HTML or displays the HTML inside the
 * tinyMCE rich text editor.
 * @export
 * @class CustomTextEditor
 * @extends {React.Component<ICustomTextEditorProps, ICustomTextEditorState>}
 */

enum Colors {
    'factbox' = '#efefef',
    'changelog' = '#fff2e0',
    'aside' = '#d8e2e6',
    'other' = '#eff6fc',
    'none' = 'rgba(255,255,255,0)',
}

export default class CustomTextEditor extends React.Component<ICustomTextEditorProps, ICustomTextEditorState> {

    /**
     * Creates an instance of CustomTextEditor.
     * Initializes the local version of tinymce.
     * @param {ICustomTextEditorProps} props
     * @memberof CustomTextEditor
     */
    public constructor(props: ICustomTextEditorProps) {
        super(props);
        this.state = {
            content: this.props.content,
            isCollapsed: true
        } as ICustomTextEditorState;

        this.handleChange = this.handleChange.bind(this);
        this.findHandler = this.findHandler.bind(this);
        this.header = this.header.bind(this);
        this.toggle = this.toggle.bind(this);
        this.keyToggle = this.keyToggle.bind(this);
    }

    /**
     *
     *
     * @memberof CustomTextEditor
     */
    public async componentDidMount() {
        if (!this.props.isReadMode) {
            this.loadEditor();
        }
    }

    /**
     *
     *
     * @private
     * @memberof CustomTextEditor
     */
    private loadEditor() {
        this.setState({ editor: undefined }, async () => {
            let loader = await import(
                /* webpackChunkName: 'tinymce' */
                './TinymceLoader');
            loader.TinymceLoader.init();
            const editor = await import(
                /* webpackChunkName: 'tinymce' */
                '@tinymce/tinymce-react');
            this.setState({ editor });
        });
    }

    /**
     *
     *
     * @memberof CustomTextEditor
     */
    public async componentDidUpdate(_prevProps: ICustomTextEditorProps, _prevState: ICustomTextEditorState) {
        if(
            _prevProps.isReadMode && !this.props.isReadMode
            || _prevProps.textBoxStyle !== this.props.textBoxStyle
            || _prevProps.backgroundColorChoice !== this.props.backgroundColorChoice
            || _prevProps.themeVariant !== this.props.themeVariant
        ) {
            this.loadEditor();
        }
    }
    /**
     * Renders the editor in read mode or edit mode depending
     * on the site page.
     * @returns {React.ReactElement<ICustomTextEditorProps>}
     * @memberof CustomTextEditor
     */
    public render(): React.ReactElement<ICustomTextEditorProps> {
        const { semanticColors } = this.props.themeVariant;
        return (
            <div style={{
                backgroundColor: semanticColors.bodyBackground,
                ...this.props.textBoxStyle === TextBoxStyle.Accordion && this.props.borderBottomChoice !== false
                ? {borderBottom: `1px solid ${semanticColors.bodyText}`}
                : {},
            }}>
                {this.props.isReadMode
                    ? this.renderReadMode()
                    : this.renderEditMode()
                }
            </div>
        );
    }

    /**
     * Displays the Editor in read mode, for all users.
     * @private
     * @returns {React.ReactElement<ICustomTextEditorProps>}
     * @memberof CustomTextEditor
     */
    private renderEditMode(): React.ReactElement<ICustomTextEditorProps> {
        let editorComponent = null;
        const { fonts } = this.props.themeVariant;
        const colors = this.getBackgroundAndTextColor();
        if (this.state.editor) {
            const runtimePath: string = require('./dummy.png');
            const skinUrl = runtimePath.substr(0, runtimePath.lastIndexOf("/"));
            editorComponent = <this.state.editor.Editor
                init={{
                    plugins: ['paste', 'link', 'image', 'lists', 'advlist', 'table'],
                    content_style: `
                        .mce-content-body {
                            background-color: ${colors.body.backgroundColor};
                            color: ${colors.body.color};
                            font-size: ${fonts.medium.fontSize};
                            font-family: ${fonts.medium.fontFamily};
                            font-weight: ${fonts.medium.fontWeight};
                            -webkit-font-smoothing: ${fonts.medium.WebkitFontSmoothing};
                            -moz-osx-font-smoothing: ${fonts.medium.MozOsxFontSmoothing};
                        }
                        .mce-content-body a {
                            color: ${colors.links.color};
                        }
                        .mce-content-body h1,
                        .mce-content-body h2 {
                            font-size: ${fonts.xxLarge.fontSize};
                            font-family: ${fonts.xxLarge.fontFamily};
                            font-weight: ${fonts.xxLarge.fontWeight};
                            -webkit-font-smoothing: ${fonts.xxLarge.WebkitFontSmoothing};
                            -moz-osx-font-smoothing: ${fonts.xxLarge.MozOsxFontSmoothing};
                        }
                        .mce-content-body h3 {
                            font-size: ${fonts['xLargePlus'].fontSize};
                            font-family: ${fonts['xLargePlus'].fontFamily};
                            font-weight: ${fonts['xLargePlus'].fontWeight};
                            -webkit-font-smoothing: ${fonts['xLargePlus'].WebkitFontSmoothing};
                            -moz-osx-font-smoothing: ${fonts['xLargePlus'].MozOsxFontSmoothing};
                        }
                        .mce-content-body h4,
                        .mce-content-body h5,
                        .mce-content-body h6 {
                            font-size: ${fonts.xLarge.fontSize};
                            font-family: ${fonts.xLarge.fontFamily};
                            font-weight: ${fonts.xLarge.fontWeight};
                            -webkit-font-smoothing: ${fonts.xLarge.WebkitFontSmoothing};
                            -moz-osx-font-smoothing: ${fonts.xLarge.MozOsxFontSmoothing};
                        }
                        .mce-content-body p,
                        .mce-content-body li {
                            font-size: ${fonts.large.fontSize};
                            font-family: ${fonts.large.fontFamily};
                            font-weight: ${fonts.large.fontWeight};
                            -webkit-font-smoothing: ${fonts.large.WebkitFontSmoothing};
                            -moz-osx-font-smoothing: ${fonts.large.MozOsxFontSmoothing};
                         }
                    `,
                    style_formats: [
                        { title: 'Overskrifter', items: [
                            { title: 'Overskrift 1', format: 'h2' },
                            { title: 'Overskrift 2', format: 'h3' },
                            { title: 'Overskrift 3', format: 'h4' },
                            { title: 'Overskrift 4', format: 'h5' },
                        ]},
                        { title: 'Brødtekst', items: [
                            { title: 'Avsnitt', format: 'p' },
                            { title: 'Sitat', format: 'blockquote' },
                            { title: 'Uformatert', format: 'pre' }
                        ]},
                        { title: 'Tekststiler', items: [
                            { title: 'Uthevet', format: 'bold' },
                            { title: 'Kursiv', format: 'italic' },
                            { title: 'Understreket', format: 'underline' },
                            { title: 'Gjennomstreket', format: 'strikethrough' },
                            { title: 'Superscript', format: 'superscript' },
                            { title: 'Subscript', format: 'subscript' },
                            { title: 'Code', format: 'code' }
                        ]},
                        { title: 'Juster', items: [
                            { title: 'Venstre', format: 'alignleft' },
                            { title: 'Midtstilt', format: 'aligncenter' },
                            { title: 'Høyre', format: 'alignright' },
                            { title: 'Fulljustert', format: 'alignjustify' }
                        ]}
                    ],
                    skin_url: skinUrl,
                    height: 400,
                    menubar: 'edit insert format, table',
                    table_default_styles: {
                        'border-collapse': 'collapse',
                        'width': '100%'
                    },
                    table_responsive_width: true,
                    convert_urls: false,
                    relative_urls: false
                }}
                initialValue={this.state.content}
                onChange={(event) => { this.handleChange(event.target.getContent()); }}
            />;
        }
        return (
            <div>
                <WebPartTitle
                    displayMode={this.props.displayMode}
                    title={this.props.title}
                    updateProperty={this.props.setTitle}
                    themeVariant={this.props.themeVariant}
                    className={styles.customTextTitle__edit}
                />
                {editorComponent}
            </div>
        );
    }

    /**
     * Displays the editor in edit mode for power users
     * who have access to click the edit button on the site page.
     * @private
     * @returns {React.ReactElement<any>}
     * @memberof CustomTextEditor
     */
    private renderReadMode(): React.ReactElement<any> {
        const {fonts} = this.props.themeVariant;

        let content = this.state.content;
        let dataInterception = /<a/gi;
        if (content) {
            content = content.replace(dataInterception, '<a data-interception="off"');
        }

        if (this.state.isCollapsed && this.props.textBoxStyle === TextBoxStyle.Accordion) {
            document.addEventListener("keydown", this.findHandler, true);
        } else {
            document.removeEventListener("keydown", this.findHandler, true);
        }

        let bodyText = (!this.state.isCollapsed && this.props.textBoxStyle === TextBoxStyle.Accordion)
            || (this.props.textBoxStyle !== TextBoxStyle.Accordion) ? (
                <article style={{
                    position: 'relative',
                    overflow: 'hidden',
                    width: '100%',
                    fontSize: fonts.large.fontSize,
                    lineHeight: 1.4,
                }}>
                    <div
                        style={{
                            ...this.getBackgroundAndTextColor().body,
                            ...this.props.textBoxStyle === TextBoxStyle.WithBackgroundColor ? {padding: '5px 8px 5px 15px'} : {},
                        }}
                        className={this.props.themeVariant.isInverted ? styles.body__inverted : styles.body}
                    >
                        {ReactHtmlParser(content)}
                    </div>
                </article>
            ) : null;

        return (
            <>
                <this.header />
                {bodyText}
            </>
        );
    }

    /**
     * Sets the state of the current TSX file and
     * invokes the saveRteContent callback with
     * the states content.
     * @private
  * @param {string} content
      * @memberof CustomTextEditor
      */
    private handleChange(content: string): void {
        this.setState({ content: content }, () => {
            this.props.saveRteContent(content);
        });
    }
    private toggle() {
        (this.state.isCollapsed) ? this.setState({ isCollapsed: false }) : this.setState({ isCollapsed: true });
    }

    private keyToggle(event) {
        if (event.keyCode !== 13) return;
        this.toggle();
    }

    private findHandler(e) {
        if ((e.keyCode == 70 && (e.ctrlKey || e.metaKey)) ||
            (e.keyCode == 191)) {
            this.toggle();
        }
    }

    private header(): JSX.Element {
        const {semanticColors} = this.props.themeVariant;
        return (
            this.props.textBoxStyle === TextBoxStyle.Accordion
            ? <button
                style={{...this.getBackgroundAndTextColor().body,
                    ...{
                        border: 'none',
                        paddingTop: '8px',
                        paddingLeft: '10px',
                        cursor: 'pointer',
                        display: 'flex',
                        margin: '0px',
                        width: '100%',
                        textAlign: 'left',
                        outline: 0,
                        transition: '.4s',
                        padding: '0px',
                        minHeight: 'inherit',
                    },
                    ...!this.state.isCollapsed ? {cursor: 'pointer'} : {},
                }}
                aria-label={this.props.title}
                aria-expanded={!this.state.isCollapsed}
                onClick={this.toggle}
                >
                <div style={{
                    fontSize: '20px',
                    paddingRight: '12px',
                    paddingLeft: '4px',
                    paddingTop: '3px',
                }}>
                    <Icon iconName={(!this.state.isCollapsed) ? 'ChevronUp' : 'ChevronDown'} />
                </div>
                <div>
                    <WebPartTitle
                        displayMode={this.props.displayMode}
                        title={this.props.title}
                        updateProperty={this.props.setTitle}
                        themeVariant={this.props.themeVariant}
                        className={styles.customTextTitle__accordion}
                    />
                </div>
                </button>
            : <WebPartTitle
                displayMode={this.props.displayMode}
                title={this.props.title}
                updateProperty={this.props.setTitle}
                themeVariant={this.props.themeVariant}
                className={styles.customTextTitle__box}
                />
        );
    }

    private getBackgroundAndTextColor() {
        if (
            this.props.textBoxStyle === TextBoxStyle.WithBackgroundColor
            && this.props.backgroundColorChoice
            && this.props.backgroundColorChoice !== 'none'
        ) return {
            body: {
                backgroundColor: Colors[this.props.backgroundColorChoice] || this.props.backgroundColor /* deprecated */ || 'inherit',
                color: this.props.themeVariant.palette['BodyText'],
            },
            links: {color: this.props.themeVariant.palette['Hyperlink']},
        };
        return {
            body: {
                backgroundColor: this.props.themeVariant.semanticColors.bodyBackground,
                color: this.props.themeVariant.semanticColors.bodyText,
            },
            links: {color: this.props.themeVariant.semanticColors.link},
        };
    }
}
