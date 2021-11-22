import { DisplayMode } from "@microsoft/sp-core-library";
import { TextBoxStyle } from "./TextBoxStyle";
import { IReadonlyTheme } from '@microsoft/sp-component-base';

export interface ICustomTextEditorProps {
    title: string;
    displayMode: DisplayMode;
    setTitle: (title: string) => void;
    saveRteContent(content: string): void;
    isReadMode: boolean;
    content: string;
    textBoxStyle: TextBoxStyle;
    backgroundColor: string; /* deprecated */
    backgroundColorChoice: string;
    borderBottomChoice: boolean;
    themeVariant: IReadonlyTheme;
}
