import { omit } from 'lodash'
import { TextSplitter } from 'langchain/text_splitter'
//import { ConfluencePagesLoader, ConfluencePagesLoaderParams } from '@langchain/community/document_loaders/web/confluence'
import { getCredentialData, getCredentialParam, handleEscapeCharacters } from '../../../src/utils'
import { ICommonObject, INode, INodeData, INodeParams, INodeOutputsValue } from '../../../src/Interface'

import { htmlToText } from 'html-to-text'
import { Document } from '@langchain/core/documents'
import { BaseDocumentLoader } from '@langchain/core/document_loaders/base'

/**
 * Interface representing the parameters for configuring the
 * ConfluencePagesLoader.
 */
interface ConfluencePagesLoaderParams {
    baseUrl: string
    spaceKey: string
    username?: string
    accessToken?: string
    personalAccessToken?: string
    label?: string
    limit?: number
    expand?: string
    maxRetries?: number
}

/**
 * Interface representing a Confluence page.
 */
export interface ConfluencePage {
    id: string
    title: string
    type: string
    body: {
        storage: {
            value: string
        }
    }
    status: string
    version?: {
        number: number
        when: string
        by: {
            displayName: string
        }
    }
}

/**
 * Interface representing the response from the Confluence API.
 */
interface ConfluenceAPIResponse {
    size: number
    results: ConfluencePage[]
}

/**
 * Class representing a document loader for loading pages from Confluence.
 * @example
 * ```typescript
 * const loader = new ConfluencePagesLoader({
 *   baseUrl: "https:
 *   spaceKey: "~EXAMPLE362906de5d343d49dcdbae5dEXAMPLE",
 *   username: "your-username",
 *   accessToken: "your-access-token",
 * });
 * const documents = await loader.load();
 * console.log(documents);
 * ```
 */
class ConfluencePagesLoader extends BaseDocumentLoader {
    public readonly baseUrl: string

    public readonly spaceKey: string

    public readonly username?: string

    public readonly accessToken?: string

    public readonly label?: string

    public readonly limit: number

    public readonly maxRetries: number

    /**
     * expand parameter for confluence rest api
     * description can be found at https://developer.atlassian.com/server/confluence/expansions-in-the-rest-api/
     */
    public readonly expand?: string

    public readonly personalAccessToken?: string

    constructor({
        baseUrl,
        spaceKey,
        username,
        accessToken,
        label,
        limit = 25,
        expand = 'body.storage,version',
        personalAccessToken,
        maxRetries = 5
    }: ConfluencePagesLoaderParams) {
        super()
        this.baseUrl = baseUrl
        this.spaceKey = spaceKey
        this.username = username
        this.accessToken = accessToken
        this.label = label
        this.limit = limit
        this.expand = expand
        this.personalAccessToken = personalAccessToken
        this.maxRetries = maxRetries
    }

    /**
     * Returns the authorization header for the request.
     * @returns The authorization header as a string, or undefined if no credentials were provided.
     */
    private get authorizationHeader(): string | undefined {
        if (this.personalAccessToken) {
            return `Bearer ${this.personalAccessToken}`
        } else if (this.username && this.accessToken) {
            const authToken = Buffer.from(`${this.username}:${this.accessToken}`).toString('base64')
            return `Basic ${authToken}`
        }

        return undefined
    }

    /**
     * Fetches all the pages in the specified space and converts each page to
     * a Document instance.
     * @param options the extra options of the load function
     * @param options.limit The limit parameter to overwrite the size to fetch pages.
     * @param options.start The start parameter to set inital offset to fetch pages.
     * @returns Promise resolving to an array of Document instances.
     */
    public async load(options?: { start?: number; label?: string; limit?: number }): Promise<Document[]> {
        try {
            const pages = await this.fetchAllPagesInSpace(options?.start, options?.label, options?.limit)
            return pages.map((page) => this.createDocumentFromPage(page))
        } catch (error) {
            console.error('Error:', error)
            return []
        }
    }

    /**
     * Fetches data from the Confluence API using the provided URL.
     * @param url The URL to fetch data from.
     * @returns Promise resolving to the JSON response from the API.
     */
    protected async fetchConfluenceData(url: string): Promise<ConfluenceAPIResponse> {
        let retryCounter = 0
        // eslint-disable-next-line no-constant-condition
        while (true) {
            retryCounter += 1
            try {
                const initialHeaders: HeadersInit = {
                    'Content-Type': 'application/json',
                    Accept: 'application/json'
                }

                const authHeader = this.authorizationHeader
                if (authHeader) {
                    initialHeaders.Authorization = authHeader
                }

                const response = await fetch(url, {
                    headers: initialHeaders
                })

                if (!response.ok) {
                    throw new Error(`Failed to fetch ${url} from Confluence: ${response.status}. Retrying...`)
                }

                return await response.json()
            } catch (error) {
                if (retryCounter >= this.maxRetries)
                    throw new Error(`Failed to fetch ${url} from Confluence (retry: ${retryCounter}): ${error}`)
            }
        }
    }

    /**
     * Recursively fetches all the pages in the specified space.
     * @param start The start parameter to paginate through the results.
     * @returns Promise resolving to an array of ConfluencePage objects.
     */
    private async fetchAllPagesInSpace(start = 0, label = this.label, limit = this.limit): Promise<ConfluencePage[]> {
        let url = `${this.baseUrl}/rest/api/content/search?cql=space=${this.spaceKey}+AND+type=page&limit=${limit}&start=${start}&expand=${this.expand}`
        if (label) {
            url += `AND+label=${label}`
        }

        const data = await this.fetchConfluenceData(url)

        if (data.size === 0) {
            return []
        }

        const nextPageStart = start + data.size
        const nextPageResults = await this.fetchAllPagesInSpace(nextPageStart, label, limit)

        return data.results.concat(nextPageResults)
    }

    /**
     * Creates a Document instance from a ConfluencePage object.
     * @param page The ConfluencePage object to convert.
     * @returns A Document instance.
     */
    private createDocumentFromPage(page: ConfluencePage): Document {
        const htmlContent = page.body.storage.value

        // Handle both self-closing and regular macros for attachments and view-file
        const htmlWithoutOtherMacros = htmlContent.replace(
            /<ac:structured-macro\s+ac:name="(attachments|view-file)"[^>]*(?:\/?>|>.*?<\/ac:structured-macro>)/gs,
            '[ATTACHMENT]'
        )

        // Extract and preserve code blocks with unique placeholders
        const codeBlocks: { language: string; code: string }[] = []
        const htmlWithPlaceholders = htmlWithoutOtherMacros.replace(
            /<ac:structured-macro.*?<ac:parameter ac:name="language">(.*?)<\/ac:parameter>.*?<ac:plain-text-body><!\[CDATA\[([\s\S]*?)\]\]><\/ac:plain-text-body><\/ac:structured-macro>/g,
            (_, language, code) => {
                const placeholder = `CODE_BLOCK_${codeBlocks.length}`
                codeBlocks.push({ language, code: code.trim() })
                return `\n${placeholder}\n`
            }
        )

        // Convert the HTML content to plain text
        let plainTextContent = htmlToText(htmlWithPlaceholders, {
            wordwrap: false,
            preserveNewlines: true
        })

        // Reinsert code blocks with proper markdown formatting
        codeBlocks.forEach(({ language, code }, index) => {
            const placeholder = `CODE_BLOCK_${index}`
            plainTextContent = plainTextContent.replace(placeholder, `\`\`\`${language}\n${code}\n\`\`\``)
        })

        // Remove empty lines
        const textWithoutEmptyLines = plainTextContent.replace(/^\s*[\r\n]/gm, '')

        // Rest of the method remains the same...
        return new Document({
            pageContent: textWithoutEmptyLines,
            metadata: {
                id: page.id,
                status: page.status,
                title: page.title,
                type: page.type,
                url: `${this.baseUrl}/spaces/${this.spaceKey}/pages/${page.id}`,
                version: page.version?.number,
                updated_by: page.version?.by?.displayName,
                updated_at: page.version?.when
            }
        })
    }
}

class Confluence_DocumentLoaders implements INode {
    label: string
    name: string
    version: number
    description: string
    type: string
    icon: string
    category: string
    baseClasses: string[]
    credential: INodeParams
    inputs: INodeParams[]
    outputs: INodeOutputsValue[]

    constructor() {
        this.label = 'Confluence'
        this.name = 'confluence'
        this.version = 2.0
        this.type = 'Document'
        this.icon = 'confluence.svg'
        this.category = 'Document Loaders'
        this.description = `Load data from a Confluence Document`
        this.baseClasses = [this.type]
        this.credential = {
            label: 'Connect Credential',
            name: 'credential',
            type: 'credential',
            credentialNames: ['confluenceCloudApi', 'confluenceServerDCApi']
        }
        this.inputs = [
            {
                label: 'Text Splitter',
                name: 'textSplitter',
                type: 'TextSplitter',
                optional: true
            },
            {
                label: 'Base URL',
                name: 'baseUrl',
                type: 'string',
                placeholder: 'https://example.atlassian.net/wiki'
            },
            {
                label: 'Space Key',
                name: 'spaceKey',
                type: 'string',
                placeholder: '~EXAMPLE362906de5d343d49dcdbae5dEXAMPLE',
                description:
                    'Refer to <a target="_blank" href="https://community.atlassian.com/t5/Confluence-questions/How-to-find-the-key-for-a-space/qaq-p/864760">official guide</a> on how to get Confluence Space Key'
            },
            {
                label: 'Limit',
                name: 'limit',
                type: 'number',
                default: 0,
                optional: true
            },
            {
                label: 'Label',
                name: 'label',
                type: 'string',
                placeholder: 'confluence label',
                optional: true
            },
            {
                label: 'Additional Metadata',
                name: 'metadata',
                type: 'json',
                description: 'Additional metadata to be added to the extracted documents',
                optional: true,
                additionalParams: true
            },
            {
                label: 'Omit Metadata Keys',
                name: 'omitMetadataKeys',
                type: 'string',
                rows: 4,
                description:
                    'Each document loader comes with a default set of metadata keys that are extracted from the document. You can use this field to omit some of the default metadata keys. The value should be a list of keys, seperated by comma. Use * to omit all metadata keys execept the ones you specify in the Additional Metadata field',
                placeholder: 'key1, key2, key3.nestedKey1',
                optional: true,
                additionalParams: true
            }
        ]
        this.outputs = [
            {
                label: 'Document',
                name: 'document',
                description: 'Array of document objects containing metadata and pageContent',
                baseClasses: [...this.baseClasses, 'json']
            },
            {
                label: 'Text',
                name: 'text',
                description: 'Concatenated string from pageContent of documents',
                baseClasses: ['string', 'json']
            }
        ]
    }

    async init(nodeData: INodeData, _: string, options: ICommonObject): Promise<any> {
        const spaceKey = nodeData.inputs?.spaceKey as string
        const baseUrl = nodeData.inputs?.baseUrl as string
        const label = nodeData.inputs?.label as string
        const limit = nodeData.inputs?.limit as number
        const textSplitter = nodeData.inputs?.textSplitter as TextSplitter
        const metadata = nodeData.inputs?.metadata
        const _omitMetadataKeys = nodeData.inputs?.omitMetadataKeys as string
        const output = nodeData.outputs?.output as string

        let omitMetadataKeys: string[] = []
        if (_omitMetadataKeys) {
            omitMetadataKeys = _omitMetadataKeys.split(',').map((key) => key.trim())
        }

        const credentialData = await getCredentialData(nodeData.credential ?? '', options)
        const accessToken = getCredentialParam('accessToken', credentialData, nodeData)
        const personalAccessToken = getCredentialParam('personalAccessToken', credentialData, nodeData)
        const username = getCredentialParam('username', credentialData, nodeData)

        let confluenceOptions: ConfluencePagesLoaderParams = {
            baseUrl,
            spaceKey,
            label,
            limit
        }

        if (accessToken) {
            // Confluence Cloud credentials
            confluenceOptions.username = username
            confluenceOptions.accessToken = accessToken
        } else if (personalAccessToken) {
            // Confluence Server/Data Center credentials
            confluenceOptions.personalAccessToken = personalAccessToken
        }

        const loader = new ConfluencePagesLoader(confluenceOptions)

        let docs = []

        if (textSplitter) {
            docs = await loader.load()
            docs = await textSplitter.splitDocuments(docs)
        } else {
            docs = await loader.load()
        }

        if (metadata) {
            const parsedMetadata = typeof metadata === 'object' ? metadata : JSON.parse(metadata)
            docs = docs.map((doc) => ({
                ...doc,
                metadata:
                    _omitMetadataKeys === '*'
                        ? {
                              ...parsedMetadata
                          }
                        : omit(
                              {
                                  ...doc.metadata,
                                  ...parsedMetadata
                              },
                              omitMetadataKeys
                          )
            }))
        } else {
            docs = docs.map((doc) => ({
                ...doc,
                metadata:
                    _omitMetadataKeys === '*'
                        ? {}
                        : omit(
                              {
                                  ...doc.metadata
                              },
                              omitMetadataKeys
                          )
            }))
        }

        if (output === 'document') {
            return docs
        } else {
            let finaltext = ''
            for (const doc of docs) {
                finaltext += `${doc.pageContent}\n`
            }
            return handleEscapeCharacters(finaltext, false)
        }
    }
}

module.exports = { nodeClass: Confluence_DocumentLoaders }
