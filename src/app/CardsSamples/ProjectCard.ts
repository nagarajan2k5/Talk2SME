declare module namespace {

    export interface Item {
        type: string;
        url: string;
        size: string;
        style: string;
        text: string;
        weight: string;
        wrap?: boolean;
        spacing: string;
        isSubtle?: boolean;
    }

    export interface Column {
        type: string;
        width: string;
        items: Item[];
    }

    export interface Body {
        type: string;
        text: string;
        weight: string;
        isSubtle: boolean;
        separator?: boolean;
        columns: Column[];
    }

    export interface Data {
        x: string;
    }

    export interface Action {
        type: string;
        title: string;
        data: Data;
    }

    export interface RootObject {
        $schema: string;
        version: string;
        type: string;
        body: Body[];
        actions: Action[];
    }

}

