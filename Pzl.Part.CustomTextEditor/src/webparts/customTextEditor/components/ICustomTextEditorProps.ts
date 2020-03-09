import { DisplayMode } from "@microsoft/sp-core-library";
import { TextBoxStyle } from "./TextBoxStyle";
export interface ICustomTextEditorProps {
    title: string;
    displayMode: DisplayMode;
    setTitle: (title: string) => void;
    saveRteContent(content: string): void;
    isReadMode: boolean;
    content: string;
    textBoxStyle: TextBoxStyle;
    backgroundColor: string;
    headerExpandColor: string;
    underlineLinks: boolean;
}