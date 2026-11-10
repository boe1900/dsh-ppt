export declare const TEMPLATE_DESIGN_FILE = "template-design.json";
export declare const MAX_TEMPLATE_DESIGN_BYTES: number;
export interface TemplateLayout {
    readonly file: string;
    readonly name: string;
    readonly purpose: string;
    readonly relationship: string;
    readonly capacity: string;
    readonly adaptation: string;
    readonly preserve: string;
    readonly replace: string;
}
export interface TemplateDesign {
    readonly version: 1;
    readonly purpose: string;
    readonly rules: {
        readonly palette: string;
        readonly typography: string;
        readonly spacing: string;
        readonly imagery: string;
        readonly data: string;
        readonly overflow: string;
    };
    readonly layouts: readonly TemplateLayout[];
}
export declare function parseTemplateDesign(bytes: Uint8Array, pages?: readonly { readonly file: string }[]): TemplateDesign;
export declare function readTemplateDesign(directory: string, pages: readonly { readonly file: string }[], required?: boolean): Promise<{ readonly bytes: Buffer; readonly design: TemplateDesign } | undefined>;
export declare function templateLayoutGuidance(layout: TemplateLayout): string;
export declare function templateDesignGuidance(design: TemplateDesign): string;
