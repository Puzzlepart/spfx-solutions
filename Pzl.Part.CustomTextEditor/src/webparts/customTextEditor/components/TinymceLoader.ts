import * as tinymce from 'tinymce';
/*
Import the plugins that you want to use here.
*/
import 'tinymce/themes/modern/theme';
import 'tinymce/plugins/paste';
import 'tinymce/plugins/link';
import 'tinymce/plugins/image';
import 'tinymce/plugins/lists';
import 'tinymce/plugins/advlist';
import 'tinymce/plugins/table';

export class TinymceLoader {
    public static init(){
        tinymce.init({});
    }
}