import * as React from 'react';
import styles from './CustomTextEditor.module.scss';
import { ICustomTextEditorProps } from './ICustomTextEditorProps';
import { ICustomTextEditorState } from './ICustomTextEditorState';
import ReactHtmlParser from 'react-html-parser';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import * as strings from 'CustomTextEditorWebPartStrings';

export enum TextBoxStyle {
    WithBackgroundColor,
    Accordion,
    Regular,
    RegularFade
}

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
export default class CustomTextEditor extends React.Component<ICustomTextEditorProps, ICustomTextEditorState> {

    /**
     * Creates an instance of CustomTextEditor.
     * Initializes the local version of tinymce.
     * @param {ICustomTextEditorProps} props
     * @memberof CustomTextEditor
     */
    public constructor(props: ICustomTextEditorProps) {
        super(props);
        //tinymce.init({});
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
            await this.loadEditor();
        }
    }

    /**
     *
     *
     * @private
     * @memberof CustomTextEditor
     */
    private async loadEditor() {
        let loader = await import(
            /* webpackChunkName: 'tinymce' */
            './TinymceLoader');
        loader.TinymceLoader.init();
        const editor = await import(
            /* webpackChunkName: 'tinymce' */
            '@tinymce/tinymce-react');
        this.setState({ editor: editor });
    }

    
    /**
     *
     *
     * @memberof CustomTextEditor
     */
    public async componentDidUpdate(_prevProps: ICustomTextEditorProps, _prevState: ICustomTextEditorState) {
        if(_prevProps.isReadMode  && !this.props.isReadMode) {
            await this.loadEditor();
        }
    }
    /**
     * Renders the editor in read mode or edit mode depending
     * on the site page.
     * @returns {React.ReactElement<ICustomTextEditorProps>}
     * @memberof CustomTextEditor
     */
    public render(): React.ReactElement<ICustomTextEditorProps> {
        return (
            <div>
                {
                    this.props.isReadMode
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
        if (this.state.editor) {
            const runtimePath: string = require('./dummy.png');
            const skinUrl = runtimePath.substr(0, runtimePath.lastIndexOf("/"));
            editorComponent = <this.state.editor.Editor
                init={{
                    plugins: ['paste', 'link', 'image', 'lists', 'advlist', 'table'],
                    content_style: `#tinymce .mce-content-body { color: #333333; font-family: "Segoe UI Web (West European)", "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif !important;}
.mce-content-body h1 {font-size: 28px; font-weight:normal;} .mce-content-body h2 { font-size: 24px; font-weight:normal;}.mce-content-body h3 {font-size: 21px; font-weight:normal;}.mce-content-body h4,
.mce-content-body h5, .mce-content-body h6 {font-size: 17px;font-weight: bold;}.mce-content-body p { font-size: 17px;} .mce-content-body li { font-size: 17px; font-weight: 300;}`,
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
                <WebPartTitle displayMode={this.props.displayMode}
                    title={this.props.title}
                    updateProperty={this.props.updateProperty} />
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
        let bgColor = null;
        if (this.props.textBoxStyle === TextBoxStyle.WithBackgroundColor) {
            bgColor = {
                backgroundColor: this.props.backgroundColor
            };
        }
        let readMore = null;
        let showFade = false;
        if (this.props.textBoxStyle === TextBoxStyle.RegularFade && this.state.isCollapsed) {
            readMore = <div className={styles.fadeMore} tabIndex={0} onKeyDown={this.keyToggle} onClick={this.toggle}>{strings.Show} <Icon iconName={'ChevronDown'} className={styles.showMoreIcon} /></div>;
            showFade = true;
        }

        let slideInStyle = this.props.textBoxStyle !== TextBoxStyle.Regular ? styles.slideInAndFade : "";

        let content = this.state.content;
        let dataInterception = /<a/gi;
        if (content) {
            content = content.replace(dataInterception, '<a data-interception="off"');
        }

        if (showFade || (this.state.isCollapsed && this.props.textBoxStyle === TextBoxStyle.Accordion)) {
            document.addEventListener("keydown", this.findHandler, true);

        } else {
            document.removeEventListener("keydown", this.findHandler, true);
        }

        let bodyText = (!this.state.isCollapsed && this.props.textBoxStyle === TextBoxStyle.Accordion)
            || (this.props.textBoxStyle !== TextBoxStyle.Accordion) ? (
                <article className={`${styles.contentBody} ${slideInStyle}`}>
                    <div style={bgColor} className={`${(this.props.textBoxStyle === TextBoxStyle.WithBackgroundColor) ? styles.backgroundColorPadding : ""} ${slideInStyle}`}>
                        {ReactHtmlParser(content)}
                    </div>
                </article>
            ) : null;

        let linkClass = this.props.underlineLinks ? styles.underline : '';

        return (
            <div className={`${styles.editor} ${linkClass}`}>
                <this.header />
                {bodyText}
                {readMore}
            </div>
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
        let bgColor = null;
        if (this.props.textBoxStyle === TextBoxStyle.Accordion) {
            if (!this.state.isCollapsed) {
                bgColor = {
                    backgroundColor: this.props.headerExpandColor
                };
            }
        }

        return (
            this.props.textBoxStyle === TextBoxStyle.Accordion ?
                <button style={bgColor} aria-label={this.props.title} aria-expanded={!this.state.isCollapsed} className={`${styles.headerContainer} ${styles.accordionContainer} ${(!this.state.isCollapsed) ? styles.expanded : ""}`} onClick={this.toggle}>
                    <div className={styles.chevron}>
                        <Icon iconName={(!this.state.isCollapsed) ? 'ChevronDown' : 'ChevronRightMed'} />
                    </div>
                    <div>
                        <WebPartTitle displayMode={this.props.displayMode} title={this.props.title} updateProperty={this.props.updateProperty} />
                    </div>
                </button> : <WebPartTitle displayMode={this.props.displayMode} title={this.props.title} updateProperty={this.props.updateProperty} />
        );
    }
}
