import { TextBoxStyle } from "./components/TextBoxStyle";

export interface ICustomTextEditorWebPartProps {
    title: string;
    Content: string;
    searchableContent: string;
    textBoxStyle: TextBoxStyle;
    backgroundColor: string; /* deprecated */
    backgroundColorChoice: string;
    borderBottomChoice: boolean;
}
