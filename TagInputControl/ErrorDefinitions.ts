/**
 * Extends the custom error to provide a better error handling capability
 */
export class CustomError extends Error {

    /**
     * Default CustomError constructor. It changes its type and captures its stack trace based on this idea: https://stackoverflow.com/a/42755876.
     * @param message the message that will be displayed to the user.
     */
    constructor(message: string) {
        super(message)
        this.name = this.constructor.name;
        this.message = message;
        if (typeof Error.captureStackTrace === 'function') {
            Error.captureStackTrace(this, this.constructor)
        } else {
            this.stack = (new Error(message)).stack
        }
        Object.setPrototypeOf(this, CustomError.prototype);
    }
}

/**
 * This extension will keep the stack trace
 */
export class TagManagerServiceError extends CustomError {
    public original: Error;

    /**
     * Default TagManagerServiceError constructor. It changes its type and builds its stack trace based on this idea: https://stackoverflow.com/a/42755876
     * @param message the message that will be displayed to the user.
     * @param error the internal error
     */
    constructor(message: string, error: Error) {
        super(message)
        if (!error) throw new Error('TagManagerServiceError requires a message and error');
        this.original = error
        //this.new_stack = this.stack
        let message_lines = (this.message.match(/\n/g) || []).length + 1;
        if (this.stack != null) {
            this.stack = this.stack.split('\n').slice(0, message_lines + 1).join('\n') + '\n' +
                error.stack
        }
        Object.setPrototypeOf(this, TagManagerServiceError.prototype);
    }
}