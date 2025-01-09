declare module 'html-to-excel-converter' {
    interface HTMLToExcelConverterOptions {
        /** Path to Python virtual environment. Default is 'venv' in package directory */
        venvPath?: string;
        /** Maximum chunk size for processing large files in bytes. Default is 5MB */
        maxChunkSize?: number;
        /** Timeout for Python script execution in milliseconds. Default is 10 minutes */
        timeout?: number;
        /** Maximum buffer size for Python process in bytes. Default is 50MB */
        maxBuffer?: number;
    }

    export default class HTMLToExcelConverter {
        constructor(options?: HTMLToExcelConverterOptions);

        /**
         * Converts HTML content to Excel format
         * @param html The HTML content to convert
         * @returns Promise<string> Base64 encoded Excel file content
         */
        convert(html: string): Promise<string>;

        /**
         * Gets the Python command path
         * @returns Promise<string> Path to Python executable
         */
        getPythonCommand(): Promise<string>;
    }
} 