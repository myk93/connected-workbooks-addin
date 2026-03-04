import { Buffer } from "buffer";
import JSZip from "jszip";

/**
 * Namespace URIs used by Excel to store the Power Query DataMashup in customXmlParts.
 * We try each in order until we find one that has data.
 */
const PQ_NAMESPACES = [
    "http://schemas.microsoft.com/DataMashup",
    "http://schemas.microsoft.com/DataExplorer",
    "http://schemas.microsoft.com/DataMashup/Temp",
    "http://schemas.microsoft.com/DataExplorer/Temp",
];

/**
 * Reads the workbook's DataMashup custom XML part and extracts the raw M expression
 * for the given query name. Returns undefined if not found or not supported.
 *
 * The DataMashup binary layout:
 *   [4 bytes: version]
 *   [4 bytes LE: packageSize]
 *   [packageSize bytes: OPC zip containing Formulas/Section1.m]
 *   [...: permissions + metadata (ignored)]
 */
export async function extractMFormula(
    context: Excel.RequestContext,
    queryName: string
): Promise<string | undefined> {
    for (const ns of PQ_NAMESPACES) {
        try {
            const parts = context.workbook.customXmlParts.getByNamespace(ns);
            parts.load("items/id");
            await context.sync();

            if (parts.items.length === 0) continue;

            const xmlResult = parts.items[0].getXml();
            await context.sync();

            const formula = await parseMFormula(xmlResult.value, queryName);
            if (formula !== undefined) return formula;
        } catch {
            continue;
        }
    }
    return undefined;
}

async function parseMFormula(dataMashupXml: string, queryName: string): Promise<string | undefined> {
    // The DataMashup XML looks like:
    //   <?xml ...?><DataMashup xmlns="...">BASE64_BLOB</DataMashup>
    const match = dataMashupXml.match(/<DataMashup[^>]*>\s*([A-Za-z0-9+/=\s]+)\s*<\/DataMashup>/);
    if (!match) return undefined;

    const base64Str = match[1].replace(/\s/g, "");
    const bytes = new Uint8Array(Buffer.from(base64Str, "base64").buffer);

    if (bytes.length < 8) return undefined;

    // Read packageSize as int32 little-endian at offset 4
    const packageSize = bytes[4] | (bytes[5] << 8) | (bytes[6] << 16) | (bytes[7] << 24);
    if (bytes.length < 8 + packageSize || packageSize <= 0) return undefined;

    const packageOPC = bytes.slice(8, 8 + packageSize);

    try {
        const zip = await JSZip.loadAsync(packageOPC);
        const section1m = await zip.file("Formulas/Section1.m")?.async("text");
        if (!section1m) return undefined;
        return extractQueryExpression(section1m, queryName);
    } catch {
        return undefined;
    }
}

/**
 * Extracts the M expression body for `queryName` from a Section1.m document.
 *
 * Section1.m looks like:
 *   section Section1;
 *   shared #"My Query" = let ... in ...;
 *   shared Query2 = ...;
 *
 * Returns just the expression without the trailing semicolon.
 */
function extractQueryExpression(section: string, queryName: string): string | undefined {
    const escaped = queryName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

    // Try quoted form first (#"Name"), then bare identifier
    const patterns = [
        new RegExp(`shared\\s+#"${escaped}"\\s*=\\s*([\\s\\S]+?)(?=\\n\\s*shared\\s+|$)`),
        new RegExp(`shared\\s+${escaped}\\s*=\\s*([\\s\\S]+?)(?=\\n\\s*shared\\s+|$)`),
    ];

    for (const pattern of patterns) {
        const m = section.match(pattern);
        if (m) {
            return m[1].trim().replace(/;\s*$/, "").trim();
        }
    }
    return undefined;
}
