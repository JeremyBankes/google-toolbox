import { Data } from '@jeremy-bankes/toolbox';
import { google, sheets_v4 } from 'googleapis';
import GoogleClient from './GoogleClient.js';

export interface GoogleSpreadsheetOptions {
    client: GoogleClient,
    id: string,
}

export interface GoogleAnchorOptions {
    row?: number,
    column?: number,
    a1?: string
}

export interface GoogleSheetRangeOptions {
    sheetTitle?: string,
    firstAnchor?: GoogleAnchor,
    secondAnchor?: GoogleAnchor,
    a1?: string
}

export class GoogleAnchor {

    public row: number | null;
    public column: number | null;

    public constructor(options: GoogleAnchorOptions) {
        if (Data.has(options, 'a1')) {
            const match = options.a1.match(/^(?:([A-Z]+[0-9]+)(?::([A-Z]+[0-9]+)?)?)$/i);
            this.row = Data.get(match, '1');
            this.column = Data.get(match, '0');
        } else {
            this.row = Data.get(options, 'row');
            this.column = Data.get(options, 'column');
        }
    }

    get a1() {
        if (this.row === null && this.column === null) {
            return '';
        } else if (this.row === null) {
            return GoogleAnchor.getColumnLettering(this.column);
        } else if (this.column === null) {
            return this.row.toString();
        } else {
            return GoogleAnchor.getColumnLettering(this.column) + this.row.toString();
        }
    }

    public static getColumnLettering(number: number) {
        const sequence: string[] = [];
        while (number > 0) {
            const remainder = (number - 1) % 26;
            sequence.push(String.fromCharCode(65 + remainder));
            number = Math.floor((number - remainder) / 26);
        }
        return sequence.reverse().join('');
    }

    public static getColumnFromLettering(lettering: string) {
        let result = 0;
        for (let i = 0; i < lettering.length; i++) {
            result += Math.pow(26, i) * (lettering.charCodeAt(lettering.length - 1 - i) - 65 + 1);
        }
        return result;
    }

    public clone() {
        return new GoogleAnchor({ row: this.row, column: this.column });
    }

}

export class GoogleSheetRange {

    public sheetTitle: string;
    public firstAnchor: GoogleAnchor;
    public secondAnchor: GoogleAnchor;

    public constructor(options: GoogleSheetRangeOptions) {
        if (Data.has(options, 'a1')) {
            const match = options.a1.match(/^(?:'?(.+?)(?:(?<!')'(?!'))?!)?(?:([A-Z]+[0-9]+)(?::([A-Z]+[0-9]+)?)?)?$/i);
            this.sheetTitle = Data.getOrThrow(match, '1');
            this.firstAnchor = new GoogleAnchor({ a1: Data.getOrThrow(match, '2') });
            this.secondAnchor = new GoogleAnchor({ a1: Data.get(match, '3', this.firstAnchor) });
        } else {
            this.sheetTitle = Data.getOrThrow(options, 'sheetTitle');
            this.firstAnchor = Data.getOrThrow(options, 'firstAnchor');
            this.secondAnchor = Data.get(options, 'secondAnchor', this.firstAnchor.clone());
        }
    }

    public get minimumRow() { return Math.min(this.firstAnchor.row, this.secondAnchor.row); }
    public get maximumRow() { return Math.max(this.firstAnchor.row, this.secondAnchor.row); }
    public get minimumColumn() { return Math.min(this.firstAnchor.column, this.secondAnchor.column); }
    public get maximumColumn() { return Math.max(this.firstAnchor.column, this.secondAnchor.column); }
    public get rowCount() { return this.maximumRow - this.minimumRow; }
    public get columnCount() { return this.maximumColumn - this.minimumColumn; }
    public get a1() { return `${this.sheetTitle}!${this.firstAnchor.a1}:${this.secondAnchor.a1}`; }

}

export default class GoogleSpreadsheet {

    private _handle: sheets_v4.Sheets;
    private _id: string;

    public constructor(options: GoogleSpreadsheetOptions) {
        this._handle = google.sheets({ auth: options.client.handle, version: 'v4' });
        this._id = options.id;
    }

    public get id() { return this._id; }

    public async get(...ranges: GoogleSheetRange[]) {
        const googleRanges: string[] = [];
        for (const range of ranges) { googleRanges.push(range.a1); }
        const response = await this._handle.spreadsheets.values.batchGet({ spreadsheetId: this.id, ranges: googleRanges });
        const results = [];
        for (const valueRange of response.data.valueRanges) {
            const range = new GoogleSheetRange({ a1: valueRange.range });
            results.push({ range, values: valueRange.values });
        }
        return results;
    }

    private async _getInformation() {
        const response = await this._handle.spreadsheets.get({ spreadsheetId: this.id });
        return response.data;
    }

}