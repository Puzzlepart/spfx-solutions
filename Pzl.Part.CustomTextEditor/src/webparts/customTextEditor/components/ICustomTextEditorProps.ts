import { DisplayMode } from "@microsoft/sp-core-library";
import { TextBoxStyle } from "./CustomTextEditor";
export interface ICustomTextEditorProps {
    title: string;
    displayMode: DisplayMode;
    updateProperty: (value: string) => void;
    saveRteContent(content: string): void;
    isReadMode: boolean;
    content: string;
    textBoxStyle: TextBoxStyle;
    backgroundColor: string;
    headerExpandColor: string;
    underlineLinks: boolean;
}