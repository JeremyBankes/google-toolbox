import { Credentials, OAuth2Client, OAuth2ClientOptions } from 'google-auth-library';
import { Data } from '@jeremy-bankes/toolbox';
import fs from 'fs';
import path from 'path';
import readline from 'readline';

export interface GoogleClientOptions {
    clientId: string,
    clientSecret: string,
    redirectUri: string,
    scopes: string[],
    tokenCacheFilePath?: string,
    additional?: OAuth2ClientOptions
}

export default class GoogleClient {

    private _handle: OAuth2Client;
    private _tokenCacheFilePath: string | null;
    private _scopes: string[];

    public constructor(options: GoogleClientOptions) {
        this._handle = new OAuth2Client({
            clientId: options.clientId,
            clientSecret: options.clientSecret,
            redirectUri: options.redirectUri,
            ...Data.get(options, 'additional', {})
        });
        this._tokenCacheFilePath = Data.get(options, 'tokenCacheFilePath');
        this._scopes = options.scopes;
    }

    public async initialize(promptForCode: boolean = true) {
        let credentials = this.getCachedCredentials();
        if (credentials === null) {
            const authorizationLink = this.getAuthorizationLink();
            console.info(`Your credentails haven't yet been cached. Please visit the following link to sign into a Google account.\n`);
            console.info(authorizationLink);
            if (promptForCode) {
                console.info('\nPlease enter the code you received after your authorization below.\n');
                console.info(
                    '\t* Note that unless you configured a route to handle your redirect URI, this page will not load.\n' +
                    '\t* You still however, can retrieve your code from the URL parameters.\n'
                );
                const terminal = readline.createInterface({ input: process.stdin, output: process.stdout });
                const code = await new Promise<string>((resolve) => terminal.question('Enter Your Code > ', resolve));
                credentials = await this.getAndCacheCredentials(code);
                terminal.close();
                if (this._tokenCacheFilePath === null) {
                    console.info('Your Google account has been authorized for this session!');
                } else {
                    console.info('Your Google account has been authorized and your credentials have been cached!');
                }
            }
        } else {
            this._handle.setCredentials(credentials);
        }
    }

    public get handle() { return this._handle; }

    public getCachedCredentials() {
        if (this._tokenCacheFilePath !== null && fs.existsSync(this._tokenCacheFilePath)) {
            const tokenCacheFileContent = fs.readFileSync(this._tokenCacheFilePath, { encoding: 'utf8' });
            return JSON.parse(tokenCacheFileContent) as Credentials;
        } else {
            return null;
        }
    }

    public async getAndCacheCredentials(code: string) {
        const { tokens } = await this._handle.getToken(code);
        if (this._tokenCacheFilePath !== null) {
            const tokenCacheFileDirectory = path.dirname(this._tokenCacheFilePath);
            if (!fs.existsSync(tokenCacheFileDirectory)) {
                fs.mkdirSync(tokenCacheFileDirectory);
            }
            fs.writeFileSync(this._tokenCacheFilePath, JSON.stringify(tokens));
        }
        return tokens;
    }

    public getAuthorizationLink() {
        return this._handle.generateAuthUrl({ access_type: 'offline', scope: this._scopes });
    }

}