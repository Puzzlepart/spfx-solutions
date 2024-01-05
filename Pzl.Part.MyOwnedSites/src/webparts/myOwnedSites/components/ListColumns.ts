import { IColumn } from "@fluentui/react";

export const ListColumns: IColumn[] = [
    {
        key: 'displayName',
        name: 'Site',
        minWidth: 200,
        maxWidth: 300,
        isResizable: true
    },
    {
        key: 'description',
        name: 'Description',
        minWidth: 200,
        maxWidth: 300,
        isResizable: true
    },
    {
        key: 'createdDate',
        name: 'Created date',
        minWidth: 70,
        maxWidth: 100,
        isResizable: true
    },
];