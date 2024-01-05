import { IColumn } from "@fluentui/react";
import * as strings from 'MyOwnedSitesWebPartStrings';

export const ListColumns: IColumn[] = [
    {
        key: 'displayName',
        name: strings.SiteColumnName,
        minWidth: 200,
        maxWidth: 300,
        isResizable: true
    },
    {
        key: 'description',
        name: strings.DescriptionColumnName,
        minWidth: 200,
        maxWidth: 300,
        isResizable: true
    },
    {
        key: 'createdDate',
        name: strings.CreatedDateColumnName,
        minWidth: 100,
        maxWidth: 100,
        isResizable: true
    },
];