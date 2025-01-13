declare module 'html-to-excel-converter' {
    /**
     * Configuration options for the HTML to Excel converter
     */
    interface HTMLToExcelConverterOptions {
        /** Path to Python virtual environment. Default is 'venv' in package directory */
        venvPath?: string;
        /** Maximum chunk size for processing large files in bytes. Default is 5MB */
        maxChunkSize?: number;
        /** Timeout for Python script execution in milliseconds. Default is 10 minutes */
        timeout?: number;
        /** Maximum buffer size for Python process in bytes. Default is 50MB */
        maxBuffer?: number;
        /** Enable debug logging. Default is false */
        debug?: boolean;
    }

    /**
     * Response structure for conversion operations
     */
    interface ConversionResponse {
        /** Indicates if the conversion was successful */
        success: boolean;
        /** Base64 encoded Excel data or error message */
        data?: string;
        /** Error message if conversion failed */
        error?: string;
    }

    /**
     * Main converter class for HTML to Excel conversion
     */
    export default class HTMLToExcelConverter {
        /**
         * Creates a new instance of HTMLToExcelConverter
         * @param options Configuration options
         */
        constructor(options?: HTMLToExcelConverterOptions);

        /**
         * Converts HTML content to Excel file
         * @param html HTML content to convert
         * @param outputPath Optional path to save the Excel file
         * @returns Promise resolving to conversion result
         */
        convert(html: string, outputPath?: string): Promise<ConversionResponse>;

        /**
         * Converts HTML content to Excel buffer
         * @param html HTML content to convert
         * @returns Promise resolving to conversion result with base64 encoded Excel data
         */
        convertToBuffer(html: string): Promise<ConversionResponse>;

        /**
         * Gets the Python command path
         * @returns Promise resolving to Python executable path
         * @throws Error if Python is not found
         */
        getPythonCommand(): Promise<string>;

        /**
         * Validates the converter configuration
         * @returns Promise resolving to boolean indicating if configuration is valid
         * @throws Error if configuration is invalid
         */
        validateConfig(): Promise<boolean>;

        /**
         * Cleans up resources used by the converter
         * @returns Promise resolving when cleanup is complete
         */
        cleanup(): Promise<void>;
    }

    /**
     * Error types that can be thrown by the converter
     */
    export enum ConversionErrorType {
        PYTHON_NOT_FOUND = 'PYTHON_NOT_FOUND',
        INVALID_HTML = 'INVALID_HTML',
        CONVERSION_FAILED = 'CONVERSION_FAILED',
        OUTPUT_ERROR = 'OUTPUT_ERROR',
        TIMEOUT = 'TIMEOUT'
    }

    /**
     * Custom error class for conversion errors
     */
    export class ConversionError extends Error {
        constructor(
            message: string,
            public type: ConversionErrorType,
            public details?: any
        );
    }
}