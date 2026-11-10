#!/usr/bin/env node
import { checkPptdProject, loadPptdProject, renderPptdProject, resolvePptdEntry } from "./pptd.js";
import { fontFace, richRuns } from "./pptd-2VqVzr_T.js";
import { MANIFEST_FIELDS, PAGE_FIELDS, ELEMENT_FIELDS } from "./pptd-schema.js";
import { importedTextLines } from "./imported-text-layout.js";
import { preparePreviewProject } from "./raster-preview-assets.js";
import { wrapTextLines } from "./text-wrap.js";
import { layoutRichText } from "./rich-text-layout.js";
import { hasSourceLayout } from "./source-layout.js";
import { access, link, mkdir, open, readFile, realpath, rename, rm, unlink, writeFile } from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { randomUUID } from "node:crypto";
import { convertPptxToPptd } from "./pptx-converter.js";
import yaml from "js-yaml";
import sharp from "sharp";
//#region src/pptd-colors.ts
/** Deterministic series palette used when a chart series omits an explicit color. */
const PPTD_CHART_SERIES_PALETTE = [
	"#2563EB",
	"#F59E0B",
	"#10B981",
	"#EF4444",
	"#8B5CF6",
	"#06B6D4"
];
function record$3(value) {
	return typeof value === "object" && value !== null && !Array.isArray(value) ? value : void 0;
}
function themeColors(project) {
	return record$3(project.theme.colors) ?? {};
}
/** Resolve a literal or theme-referenced PPTD color without applying a silent fallback. */
function resolvePptdColor(project, value) {
	let current = value;
	const colors = themeColors(project);
	for (let depth = 0; depth < 8 && typeof current === "string" && current.startsWith("$"); depth += 1) current = colors[current.slice(1)];
	return typeof current === "string" && /^#[0-9a-f]{6}(?:[0-9a-f]{2})?$/iu.test(current) ? current : void 0;
}
/**
* Normalize chart-series colors from either color scalars or solid paint objects.
* Pie series may supply an array; unsupported paints and invalid colors return undefined.
*/
function resolvePptdChartSeriesColors(project, value) {
	if (value === void 0) return void 0;
	const values = Array.isArray(value) ? value : [value];
	if (values.length === 0) return void 0;
	const resolved = values.map((item) => {
		const paint = record$3(item);
		return resolvePptdColor(project, paint === void 0 ? item : paint.type === "solid" ? paint.color : void 0);
	});
	return resolved.every((item) => item !== void 0) ? resolved : void 0;
}
//#endregion
//#region src/pptd-preview.ts
/** Deterministic local SVG/PNG preview renderer for PPTD CLI screenshot workflows. */
function record$2(value) {
	return typeof value === "object" && value !== null && !Array.isArray(value) ? value : void 0;
}
function number$1(value) {
	return typeof value === "number" && Number.isFinite(value) ? value : void 0;
}
function string$1(value) {
	return typeof value === "string" ? value : void 0;
}
function tuple$1(value, size) {
	if (!Array.isArray(value) || value.length !== size) return void 0;
	const values = value.map(number$1);
	return values.every((item) => item !== void 0) ? values : void 0;
}
function escapeXml(value) {
	return value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll("\"", "&quot;");
}
function themeMap$1(project, key) {
	return record$2(project.theme[key]) ?? {};
}
function color(project, value, fallback = "#000000") {
	return resolvePptdColor(project, value) ?? fallback;
}
function chartSeriesColor(project, value, fallback) {
	return resolvePptdChartSeriesColors(project, value)?.[0] ?? fallback;
}
function chartSeriesPaint(series) {
	const type = string$1(series.type) ?? "bar";
	return type === "line" || type === "area" || type === "radar" ? series.lineColor ?? series.areaColor : series.fill;
}
function plainText$1(value) {
	return value.replace(/<br\s*\/?\s*>/giu, "\n").replace(/<\/p\s*>/giu, "\n").replace(/<li(?:\s[^>]*)?>/giu, "• ").replace(/<\/li\s*>/giu, "\n").replace(/<[^>]+>/gu, "").replace(/&lt;/gu, "<").replace(/&gt;/gu, ">").replace(/&amp;/gu, "&").replace(/&quot;/gu, "\"").replace(/\n{3,}/gu, "\n\n").trim();
}
function frame$1(element) {
	const values = tuple$1(element.bounds, 4) ?? [
		0,
		0,
		0,
		0
	];
	return [
		values[0] ?? 0,
		values[1] ?? 0,
		values[2] ?? 0,
		values[3] ?? 0
	];
}
function transform(element, x, y, width, height) {
	const operations = [];
	const rotation = number$1(element.rotation) ?? 0;
	if (rotation !== 0) operations.push(`rotate(${rotation} ${x + width / 2} ${y + height / 2})`);
	const flip = Array.isArray(element.flip) ? element.flip : [];
	if (flip[0] === true) operations.push(`translate(${2 * x + width} 0) scale(-1 1)`);
	if (flip[1] === true) operations.push(`translate(0 ${2 * y + height}) scale(1 -1)`);
	return operations.length === 0 ? "" : ` transform="${operations.join(" ")}"`;
}
function fillPaint(project, fillValue, definitions) {
	const fill = record$2(fillValue);
	if (fill?.type === "gradient" && Array.isArray(fill.stops) && fill.stops.length >= 2) {
		const id = definitions.next("gradient");
		const stops = fill.stops.map((rawStop) => {
			const stop = record$2(rawStop) ?? {};
			const value = color(project, stop.color);
			const alpha = value.length === 9 ? Number.parseInt(value.slice(7, 9), 16) / 255 : 1;
			return `<stop offset="${Math.round(Math.max(0, Math.min(1, number$1(stop.position) ?? 0)) * 100)}%" stop-color="${value.slice(0, 7)}" stop-opacity="${alpha}"/>`;
		}).join("");
		if (fill.gradientType === "radial") definitions.definitions.push(`<radialGradient id="${id}">${stops}</radialGradient>`);
		else definitions.definitions.push(`<linearGradient id="${id}" x1="0" y1="0.5" x2="1" y2="0.5" gradientTransform="rotate(${number$1(fill.angle) ?? 0} 0.5 0.5)">${stops}</linearGradient>`);
		return `url(#${id})`;
	}
	if (fill?.type === "solid") return color(project, fill.color);
	return "none";
}
function strokePaint(project, value) {
	const border = record$2(value);
	if (border === void 0) return "stroke=\"none\"";
	const dash = border.style === "dot" ? " stroke-dasharray=\"1 2\"" : border.style === "dash" ? " stroke-dasharray=\"5 4\"" : "";
	return `stroke="${color(project, border.color)}" stroke-width="${number$1(border.width) ?? 1}"${dash}`;
}
function renderShape$1(project, element, definitions) {
	const [x, y, width, height] = frame$1(element);
	const common = `fill="${fillPaint(project, element.fill, definitions)}" ${strokePaint(project, element.border)}${transform(element, x, y, width, height)}`;
	const shape = string$1(element.shapeName) ?? "rect";
	if (shape === "custom") {
		const viewBox = tuple$1(element.viewBox, 2) ?? [width, height];
		return `<g${transform(element,x,y,width,height)}><svg x="${x}" y="${y}" width="${width}" height="${height}" viewBox="0 0 ${viewBox[0] ?? width} ${viewBox[1] ?? height}" overflow="visible"><path d="${escapeXml(string$1(element.path) ?? "")}" fill-rule="evenodd" fill="${fillPaint(project,element.fill,definitions)}" ${strokePaint(project,element.border)}/></svg></g>`;
	}
	if (shape === "ellipse" || shape === "donut") return `<ellipse cx="${x + width / 2}" cy="${y + height / 2}" rx="${width / 2}" ry="${height / 2}" ${common}/>`;
	if (shape === "triangle") return `<path d="M ${x + width / 2} ${y} L ${x + width} ${y + height} L ${x} ${y + height} Z" ${common}/>`;
	if (shape === "diamond") return `<path d="M ${x + width / 2} ${y} L ${x + width} ${y + height / 2} L ${x + width / 2} ${y + height} L ${x} ${y + height / 2} Z" ${common}/>`;
	if (shape.includes("Arrow")) return `<path d="M ${x} ${y + height * .3} H ${x + width * .65} V ${y} L ${x + width} ${y + height / 2} L ${x + width * .65} ${y + height} V ${y + height * .7} H ${x} Z" ${common}/>`;
	return `<rect x="${x}" y="${y}" width="${width}" height="${height}" rx="${shape === "roundRect" || shape.startsWith("round") ? Math.min(width, height) * .12 : 0}" ${common}/>`;
}
function renderLine$1(project, element, definitions) {
	const [x, y, width, height] = frame$1(element);
	const viewBox = tuple$1(element.viewBox, 2) ?? [width || 1, height || 1];
	const [first = [0, 0], ...rest] = (string$1(element.points) ?? "").trim().split(/\s+/u).map((value) => value.split(",").map(Number));
	const scaleX = width / (viewBox[0] || 1);
	const scaleY = height / (viewBox[1] || 1);
	const projected = (point) => `${x + (point[0] ?? 0) * scaleX} ${y + (point[1] ?? 0) * scaleY}`;
	const pathData = rest.length === 3 ? `M ${projected(first)} C ${rest.map(projected).join(" ")}` : `M ${projected(first)} ${rest.map((point) => `L ${projected(point)}`).join(" ")}`;
	const arrow = Array.isArray(element.arrow) ? element.arrow : [];
	const marker = (kind) => {
		const id = definitions.next("arrow");
		if (kind === null || kind === void 0) return "";
		definitions.definitions.push(`<marker id="${id}" markerWidth="8" markerHeight="8" refX="7" refY="4" orient="auto"><path d="M0,0 L8,4 L0,8 Z" fill="context-stroke"/></marker>`);
		return id;
	};
	const start = marker(arrow[0]);
	const end = marker(arrow[1]);
	return `<path d="${pathData}" fill="none" ${strokePaint(project, element.border)}${start === "" ? "" : ` marker-start="url(#${start})"`}${end === "" ? "" : ` marker-end="url(#${end})"`}${transform(element, x, y, width, height)}/>`;
}
function glyphWidth$1(character, fontSize) {
	if (/\s/u.test(character)) return fontSize * .33;
	if (/^[\u0000-\u00ff]$/u.test(character)) {
		if (/[A-Z@#%&]/u.test(character)) return fontSize * .7;
		if (/[a-z0-9]/u.test(character)) return fontSize * .55;
		return fontSize * .45;
	}
	return fontSize;
}
function wrappedLines(text, width, fontSize, wrap) {
 return wrapTextLines(text, width * .95, text => Array.from(text).reduce((sum, ch) => sum + glyphWidth$1(ch, fontSize), 0), wrap);
}

function renderText$1(project, element, definitions) {
    if (hasSourceLayout(element) && Array.isArray(element.content?.paragraphs)) {
      const [x,y,w,h] = frame$1(element), content = element.content;
      const lines = importedTextLines(content,[x,y,w,h]);
      const texts = lines.map(line => `<text x="${line.x}" y="${line.baseline}" xml:space="preserve">${line.runs.map(run=>{
        const o={...content,...run.options};
        return `<tspan font-family="${escapeXml(fontFace(o.fontFamily,'Arial',run.text))}" font-size="${o.fontSize ?? 18}" font-weight="${o.bold?700:400}" font-style="${o.italic?'italic':'normal'}" fill="${color(project,o.color)}" letter-spacing="${o.letterSpacing ?? 0}" baseline-shift="${o.baseline ?? 0}">${escapeXml(run.text)}</tspan>`;
      }).join('')}</text>`).join('');
      const clip = definitions.next('imported-text');
      if(content.overflow!=='visible') definitions.definitions.push(`<clipPath id="${clip}"><rect x="${x}" y="${y}" width="${w}" height="${h}"/></clipPath>`);
      return `<g${content.overflow==='visible'?'':` clip-path="url(#${clip})"`} opacity="${element.opacity ?? 1}"${transform(element,x,y,w,h)}>${texts}</g>`;
    }

	const [x, y, width, height] = frame$1(element);
	const content = record$2(element.content) ?? {};
	const style = {
		...typeof content.style === "string" && content.style.startsWith("$") ? record$2(themeMap$1(project, "textStyles")[content.style.slice(1)]) ?? {} : {},
		...content
	};
	const align = Array.isArray(style.align) ? style.align : [];
	const anchor = align[0] === "center" ? "middle" : align[0] === "right" ? "end" : "start";
	const startX = anchor === "middle" ? x + width / 2 : anchor === "end" ? x + width : x;
	const raw = string$1(content.text) ?? "";
	const runs = /<[^>]+>/u.test(raw) ? richRuns(project, raw) : [{ text: raw, options: {} }];
	const layout = layoutRichText(runs, style, width, height);
	let top = align[1] === "middle" ? y + (height - layout.requiredHeight) / 2 : align[1] === "bottom" ? y + height - layout.requiredHeight : y;
	const family = fontFace(style.fontFamily, "Arial", raw);
	const opacity = number$1(element.opacity) ?? 1;
	const clipId = definitions.next("text-clip");
	definitions.definitions.push(`<clipPath id="${clipId}"><rect x="${x}" y="${y}" width="${Math.max(0, width)}" height="${Math.max(0, height)}"/></clipPath>`);
	const lines = layout.lines.map(line => {
		const baseline = top + line.fontSize; top += line.height;
		return `<text x="${startX}" y="${baseline}" text-anchor="${anchor}" font-family="${escapeXml(family)}" fill="${color(project, style.color)}">${line.runs.map(run => {
			const options = run.options;
			const fill = typeof options.color === "string" && /^[0-9a-f]{6}$/iu.test(options.color) ? `#${options.color}` : color(project, options.color);
			return `<tspan font-size="${options.fontSize}" font-family="${escapeXml(fontFace(options.fontFace ?? style.fontFamily, family, run.text))}" font-weight="${options.bold ? 700 : 400}" font-style="${options.italic ? "italic" : "normal"}" fill="${fill}" letter-spacing="${number$1(options.letterSpacing) ?? 0}">${escapeXml(run.text)}</tspan>`;
		}).join("")}</text>`;
	}).join("");
	return `<g clip-path="url(#${clipId})"><g opacity="${opacity}"${transform(element, x, y, width, height)}>${lines}</g></g>`;
}
function renderImage$1(project, element, definitions) {
	const [x, y, width, height] = frame$1(element);
	const asset = typeof element.src === "string" ? project.source.assets.get(element.src.replace(/^\.\//u, "")) : void 0;
	if (asset === void 0) return "";
	const fit = string$1(record$2(element.fit)?.mode) ?? "cover";
	const preserve = fit === "contain" ? "xMidYMid meet" : fit === "fill" ? "none" : "xMidYMid slice";
	const clipId = definitions.next("image-clip");
	const cropShape = record$2(element.cropShape);
	const clip = cropShape?.shapeName === "ellipse" ? `<ellipse cx="${x + width / 2}" cy="${y + height / 2}" rx="${width / 2}" ry="${height / 2}"/>` : `<rect x="${x}" y="${y}" width="${width}" height="${height}" rx="${cropShape?.shapeName === "roundRect" ? Math.min(width, height) * .12 : 0}"/>`;
	definitions.definitions.push(`<clipPath id="${clipId}">${clip}</clipPath>`);
	return `<image x="${x}" y="${y}" width="${width}" height="${height}" href="${`data:${asset.mediaType};base64,${Buffer.from(asset.bytes).toString("base64")}`}" preserveAspectRatio="${preserve}" clip-path="url(#${clipId})" opacity="${number$1(element.opacity) ?? 1}"${transform(element, x, y, width, height)}/>`;
}
function renderTable$1(project, element) {
	const [x, y, width, height] = frame$1(element);
	const columns = Array.isArray(element.columnWidths) ? element.columnWidths.map((value) => number$1(value) ?? 0) : [];
	const rowHeights = Array.isArray(element.rowHeights) ? element.rowHeights.map((value) => number$1(value) ?? 0) : [];
	const rows = Array.isArray(element.rows) ? element.rows : [];
	const occupied = Array.from({ length: rows.length }, () => Array.from({ length: columns.length }, () => false));
	const parts = [];
	for (const [rowIndex, rawRow] of rows.entries()) {
		if (!Array.isArray(rawRow)) continue;
		let columnIndex = 0;
		for (const rawCell of rawRow) {
			while (occupied[rowIndex]?.[columnIndex] === true) columnIndex += 1;
			const cell = record$2(rawCell) ?? { text: String(rawCell ?? "") };
			const colSpan = number$1(cell.colSpan) ?? 1;
			const rowSpan = number$1(cell.rowSpan) ?? 1;
			const cellX = x + columns.slice(0, columnIndex).reduce((sum, value) => sum + value, 0) * width;
			const cellY = y + rowHeights.slice(0, rowIndex).reduce((sum, value) => sum + value, 0) * height;
			const cellWidth = columns.slice(columnIndex, columnIndex + colSpan).reduce((sum, value) => sum + value, 0) * width;
			const cellHeight = rowHeights.slice(rowIndex, rowIndex + rowSpan).reduce((sum, value) => sum + value, 0) * height;
			for (let row = rowIndex; row < rowIndex + rowSpan; row += 1) {
				const occupiedRow = occupied[row];
				if (occupiedRow === void 0) continue;
				for (let column = columnIndex; column < columnIndex + colSpan; column += 1) occupiedRow[column] = true;
			}
			const fill = fillPaint(project, cell.fill, {
				definitions: [],
				next: () => ""
			});
			parts.push(`<rect x="${cellX}" y="${cellY}" width="${cellWidth}" height="${cellHeight}" fill="${fill === "none" ? "#FFFFFF" : fill}" ${strokePaint(project, cell.border ?? {
				color: "#D1D5DB",
				width: 1
			})}/>`);
			parts.push(`<text x="${cellX + cellWidth / 2}" y="${cellY + cellHeight / 2}" text-anchor="middle" dominant-baseline="middle" font-family="${escapeXml(fontFace(cell.fontFamily, "Arial", string$1(cell.text) ?? ""))}" font-size="${number$1(cell.fontSize) ?? 11}" fill="${color(project, cell.color)}">${escapeXml(plainText$1(string$1(cell.text) ?? ""))}</text>`);
			columnIndex += colSpan;
		}
	}
	return parts.join("");
}
function renderChart$1(project, element) {
	const [x, y, width, height] = frame$1(element);
	const data = record$2(element.data) ?? {};
	const columns = Array.isArray(data.cols) ? data.cols.map(String) : [];
	const rows = Array.isArray(data.rows) ? data.rows : [];
	const seriesDefaults = record$2(element.seriesDefaults) ?? {};
	const series = (Array.isArray(element.series) ? element.series.map(record$2).filter((item) => item !== void 0) : []).map((item) => ({
		...record$2(seriesDefaults[string$1(item.type) ?? ""]) ?? {},
		...item
	}));
	const colors = PPTD_CHART_SERIES_PALETTE;
	const firstEncode = record$2(series[0]?.encode) ?? {};
	const firstXIndex = columns.indexOf(string$1(firstEncode.x) ?? "");
	const firstYIndex = columns.indexOf(string$1(firstEncode.y) ?? "");
	const firstXValues = rows.map((row) => row[firstXIndex]);
	const firstYValues = rows.map((row) => row[firstYIndex]);
	const firstXNumeric = firstXIndex >= 0 && firstXValues.every((value) => Number.isFinite(Number(value)));
	const firstYNumeric = firstYIndex >= 0 && firstYValues.every((value) => Number.isFinite(Number(value)));
	const horizontalBars = string$1(series[0]?.type) === "bar" && firstXNumeric && !firstYNumeric;
	const padding = {
		left: horizontalBars ? 68 : 42,
		right: 24,
		top: 16,
		bottom: 36
	};
	const plotX = x + padding.left;
	const plotY = y + padding.top;
	const plotW = Math.max(1, width - padding.left - padding.right);
	const plotH = Math.max(1, height - padding.top - padding.bottom);
	const parts = [`<rect x="${x}" y="${y}" width="${width}" height="${height}" fill="none" stroke="#D1D5DB"/>`];
	if ((string$1(series[0]?.type) ?? "bar") === "pie") {
		const encode = record$2(series[0]?.encode) ?? {};
		const valueIndex = columns.indexOf(string$1(encode.value) ?? columns[1] ?? "");
		const categoryIndex = columns.indexOf(string$1(encode.category) ?? columns[0] ?? "");
		const values = rows.map((row) => Math.max(0, Number(row[valueIndex] ?? 0)));
		const total = values.reduce((sum, value) => sum + value, 0) || 1;
		const cx = plotX + plotW / 2;
		const cy = plotY + plotH / 2;
		const radius = Math.min(plotW, plotH) * .38;
		let angle = -Math.PI / 2;
		const configuredColors = resolvePptdChartSeriesColors(project, series[0]?.fill) ?? [];
		values.forEach((value, index) => {
			const next = angle + value / total * Math.PI * 2;
			const large = next - angle > Math.PI ? 1 : 0;
			const x1 = cx + Math.cos(angle) * radius;
			const y1 = cy + Math.sin(angle) * radius;
			const x2 = cx + Math.cos(next) * radius;
			const y2 = cy + Math.sin(next) * radius;
			const sliceColor = (configuredColors.length === 0 ? void 0 : configuredColors[index % configuredColors.length]) ?? colors[index % colors.length] ?? "#2563EB";
			parts.push(`<path d="M ${cx} ${cy} L ${x1} ${y1} A ${radius} ${radius} 0 ${large} 1 ${x2} ${y2} Z" fill="${sliceColor}"/>`);
			const labelAngle = (angle + next) / 2;
			parts.push(`<text x="${cx + Math.cos(labelAngle) * radius * .65}" y="${cy + Math.sin(labelAngle) * radius * .65}" text-anchor="middle" dominant-baseline="middle" font-size="10" fill="#FFFFFF">${escapeXml(String(rows[index]?.[categoryIndex] ?? ""))}</text>`);
			angle = next;
		});
		const innerRadius = number$1(series[0]?.innerRadius) ?? 0;
		if (innerRadius > 0) parts.push(`<circle cx="${cx}" cy="${cy}" r="${radius * Math.min(.95, innerRadius)}" fill="#FFFFFF"/>`);
		return parts.join("");
	}
	if (horizontalBars) {
		const prepared = series.map((item, index) => {
			const encode = record$2(item.encode) ?? {};
			const valueIndex = columns.indexOf(string$1(encode.x) ?? "");
			const categoryIndex = columns.indexOf(string$1(encode.y) ?? "");
			return {
				item,
				values: rows.map((row) => Number(row[valueIndex] ?? 0)),
				categories: rows.map((row) => String(row[categoryIndex] ?? "")),
				color: chartSeriesColor(project, chartSeriesPaint(item), colors[index % colors.length] ?? "#2563EB")
			};
		});
		const allValues = prepared.flatMap((item) => item.values);
		const minimum = Math.min(0, ...allValues);
		const maximum = Math.max(1, ...allValues);
		const scaleX = (value) => plotX + (value - minimum) / (maximum - minimum || 1) * plotW;
		const zeroX = scaleX(0);
		const slot = plotH / Math.max(1, rows.length);
		parts.push(`<line x1="${zeroX}" y1="${plotY}" x2="${zeroX}" y2="${plotY + plotH}" stroke="#9CA3AF"/>`);
		rows.forEach((_, rowIndex) => {
			const label = prepared[0]?.categories[rowIndex] ?? "";
			parts.push(`<text x="${plotX - 6}" y="${plotY + slot * (rowIndex + .5)}" text-anchor="end" dominant-baseline="middle" font-size="9" fill="#4B5563">${escapeXml(label)}</text>`);
		});
		for (const [seriesIndex, preparedSeries] of prepared.entries()) preparedSeries.values.forEach((value, rowIndex) => {
			const barHeight = slot * .68 / Math.max(1, prepared.length);
			const barY = plotY + slot * rowIndex + slot * .16 + seriesIndex * barHeight;
			const valueX = scaleX(value);
			parts.push(`<rect x="${Math.min(valueX, zeroX)}" y="${barY}" width="${Math.max(1, Math.abs(valueX - zeroX))}" height="${barHeight}" fill="${preparedSeries.color}"/>`);
			if (record$2(preparedSeries.item.dataLabels)?.show === true) parts.push(`<text x="${valueX + (value >= 0 ? 4 : -4)}" y="${barY + barHeight / 2}" text-anchor="${value >= 0 ? "start" : "end"}" dominant-baseline="middle" font-size="${number$1(record$2(preparedSeries.item.dataLabels)?.fontSize) ?? 9}" fill="#374151">${escapeXml(String(value))}</text>`);
		});
		if (record$2(element.legend)?.show === true && prepared.length > 1) prepared.forEach((item, index) => parts.push(`<rect x="${plotX + index * 120}" y="${y + height - 14}" width="8" height="8" fill="${item.color}"/><text x="${plotX + index * 120 + 12}" y="${y + height - 7}" font-size="8" fill="#4B5563">${escapeXml(string$1(item.item.name) ?? `Series ${index + 1}`)}</text>`));
		return parts.join("");
	}
	const allValues = [];
	const prepared = series.map((item, index) => {
		const encode = record$2(item.encode) ?? {};
		const valueColumn = string$1(encode.y) ?? string$1(encode.value) ?? columns[1] ?? "";
		const valueIndex = columns.indexOf(valueColumn);
		const values = rows.map((row) => Number(row[valueIndex] ?? 0));
		const xIndex = columns.indexOf(string$1(encode.x) ?? columns[0] ?? "");
		const xValues = rows.map((row) => row[xIndex]);
		allValues.push(...values);
		const fallbackColor = colors[index % colors.length] ?? "#2563EB";
		return {
			item,
			values,
			xValues,
			color: chartSeriesColor(project, chartSeriesPaint(item), fallbackColor)
		};
	});
	const minimum = Math.min(0, ...allValues);
	const maximum = Math.max(1, ...allValues);
	const scaleY = (value) => plotY + plotH - (value - minimum) / (maximum - minimum || 1) * plotH;
	parts.push(`<line x1="${plotX}" y1="${plotY + plotH}" x2="${plotX + plotW}" y2="${plotY + plotH}" stroke="#9CA3AF"/>`);
	for (const [seriesIndex, preparedSeries] of prepared.entries()) {
		const type = string$1(preparedSeries.item.type) ?? "bar";
		const slot = plotW / Math.max(1, rows.length);
		if (type === "line" || type === "area" || type === "scatter") {
			const numericX = preparedSeries.xValues.every((value) => Number.isFinite(Number(value)));
			const xNumbers = preparedSeries.xValues.map(Number);
			const minX = numericX ? Math.min(...xNumbers) : 0;
			const maxX = numericX ? Math.max(...xNumbers) : Math.max(1, rows.length - 1);
			const pointX = (index) => {
				if (!numericX) return plotX + slot * (index + .5);
				return plotX + ((xNumbers[index] ?? minX) - minX) / (maxX - minX || 1) * plotW;
			};
			const path = preparedSeries.values.map((value, index) => `${index === 0 ? "M" : "L"} ${pointX(index)} ${scaleY(value)}`).join(" ");
			if (type === "area") parts.push(`<path d="${path} L ${plotX + slot * (preparedSeries.values.length - .5)} ${scaleY(0)} L ${plotX + slot * .5} ${scaleY(0)} Z" fill="${preparedSeries.color}" opacity="0.3"/>`);
			if (type === "scatter") preparedSeries.values.forEach((value, index) => parts.push(`<circle cx="${pointX(index)}" cy="${scaleY(value)}" r="3" fill="${preparedSeries.color}"/>`));
			else parts.push(`<path d="${path}" fill="none" stroke="${preparedSeries.color}" stroke-width="2"/>`);
			continue;
		}
		preparedSeries.values.forEach((value, index) => {
			const barWidth = slot * .72 / Math.max(1, prepared.length);
			const barX = plotX + slot * index + slot * .14 + seriesIndex * barWidth;
			const barY = scaleY(Math.max(value, 0));
			const zeroY = scaleY(0);
			parts.push(`<rect x="${barX}" y="${Math.min(barY, zeroY)}" width="${barWidth}" height="${Math.max(1, Math.abs(zeroY - barY))}" fill="${preparedSeries.color}"/>`);
			if (record$2(preparedSeries.item.dataLabels)?.show === true) parts.push(`<text x="${barX + barWidth / 2}" y="${barY - 3}" text-anchor="middle" font-size="${number$1(record$2(preparedSeries.item.dataLabels)?.fontSize) ?? 9}" fill="#374151">${escapeXml(String(value))}</text>`);
		});
	}
	const categoryLabels = prepared[0]?.xValues ?? [];
	if (categoryLabels.some((value) => !Number.isFinite(Number(value)))) categoryLabels.forEach((value, index) => parts.push(`<text x="${plotX + plotW / Math.max(1, rows.length) * (index + .5)}" y="${plotY + plotH + 13}" text-anchor="middle" font-size="8" fill="#4B5563">${escapeXml(String(value))}</text>`));
	return parts.join("");
}
function renderElement$1(project, element, definitions) {
	if (element.elementType === "shape") return renderShape$1(project, element, definitions);
	if (element.elementType === "line") return renderLine$1(project, element, definitions);
	if (element.elementType === "text") return renderText$1(project, element, definitions);
	if (element.elementType === "image") return renderImage$1(project, element, definitions);
	if (element.elementType === "table") return renderTable$1(project, element);
	if (element.elementType === "chart") return renderChart$1(project, element);
	if (element.elementType === "icon") {
		const [x, y, width, height] = frame$1(element);
		return `<circle cx="${x + width / 2}" cy="${y + height / 2}" r="${Math.min(width, height) / 2}" fill="${color(project, record$2(element.fill)?.color)}"/><text x="${x + width / 2}" y="${y + height / 2}" text-anchor="middle" dominant-baseline="middle" font-size="${Math.min(width, height) * .35}" fill="#FFFFFF">${escapeXml(string$1(element.iconName) ?? "icon")}</text>`;
	}
	return "";
}
/** Render one PPTD page to a standalone SVG string without network access. */
function renderPptdPageSvg(project, pageIndex) {
	const page = project.pages[pageIndex];
	if (page === void 0) throw new Error(`PPTD page ${pageIndex + 1} does not exist`);
	let counter = 0;
	const definitions = {
		definitions: [],
		next(prefix) {
			counter += 1;
			return `${prefix}-${pageIndex + 1}-${counter}`;
		}
	};
	const background = fillPaint(project, page.background ?? {
		type: "solid",
		color: "#FFFFFF"
	}, definitions);
	const content = page.elements.map((element) => renderElement$1(project, element, definitions)).join("");
	return `<svg xmlns="http://www.w3.org/2000/svg" width="${project.width}" height="${project.height}" viewBox="0 0 ${project.width} ${project.height}"><defs>${definitions.definitions.join("")}</defs><rect width="${project.width}" height="${project.height}" fill="${background === "none" ? "#FFFFFF" : background}"/>${content}</svg>`;
}
/** Render one PPTD page to PNG bytes through the bundled local SVG rasterizer. */
async function renderPptdPagePng(project, pageIndex, scale = 2) {
	if (!Number.isFinite(scale) || scale <= 0 || scale > 8) throw new Error("screenshot scale must be greater than 0 and at most 8");
	const svg = renderPptdPageSvg(await preparePreviewProject(project, pageIndex, scale), pageIndex);
	// Rasterize vectors at the requested density before the final pixel-size rounding.
	const bytes = await sharp(Buffer.from(svg), { density: Math.max(1, Math.round(72 * scale)) }).resize({ width: Math.round(project.width * scale) }).png().toBuffer();
	return new Uint8Array(bytes);
}
//#endregion
//#region src/pptd-publish.ts
/** Atomic, non-overwriting publication primitives shared by PPTD CLI and model tools. */
async function pathExists(target) {
	try {
		await access(target);
		return true;
	} catch {
		return false;
	}
}
function validateDirectoryTarget(target) {
	const resolved = path.resolve(target);
	if (resolved === path.parse(resolved).root) throw new Error("Output directory cannot be a filesystem root");
	return resolved;
}
/** Atomically publish one file, preserving an existing target unless replacement was explicitly requested. */
async function publishPptdFile(target, bytes, replace) {
	const resolved = path.resolve(target);
	await mkdir(path.dirname(resolved), { recursive: true });
	const temporary = path.join(path.dirname(resolved), `.${path.basename(resolved)}.${randomUUID()}.tmp`);
	const handle = await open(temporary, "wx", 384);
	try {
		await handle.writeFile(bytes);
		await handle.sync();
	} finally {
		await handle.close();
	}
	try {
		if (replace) await rename(temporary, resolved);
		else {
			await link(temporary, resolved);
			await unlink(temporary);
		}
	} catch (error) {
		await rm(temporary, { force: true });
		if (!replace && error.code === "EEXIST") throw new Error(`Output already exists: ${resolved}`);
		throw error;
	}
}
/** Atomically publish one directory, preserving an existing target unless replacement was explicitly requested. */
async function publishPptdDirectory(target, replace, writer) {
	const resolved = validateDirectoryTarget(target);
	await mkdir(path.dirname(resolved), { recursive: true });
	const stage = path.join(path.dirname(resolved), `.${path.basename(resolved)}.${randomUUID()}.stage`);
	const backup = path.join(path.dirname(resolved), `.${path.basename(resolved)}.${randomUUID()}.backup`);
	await mkdir(stage, {
		recursive: false,
		mode: 448
	});
	try {
		await writer(stage);
		if (!replace) {
			try {
				await rename(stage, resolved);
			} catch (error) {
				if (error.code === "EEXIST" || error.code === "ENOTEMPTY") throw new Error(`Output directory already exists: ${resolved}`);
				throw error;
			}
			return;
		}
		const exists = await pathExists(resolved);
		if (exists) await rename(resolved, backup);
		try {
			await rename(stage, resolved);
			if (exists) await rm(backup, {
				recursive: true,
				force: true
			});
		} catch (error) {
			if (exists && await pathExists(backup) && !await pathExists(resolved)) await rename(backup, resolved);
			throw error;
		}
	} finally {
		await rm(stage, {
			recursive: true,
			force: true
		});
		await rm(backup, {
			recursive: true,
			force: true
		});
	}
}
function safeOutputRelative(value) {
	if (value === "" || path.isAbsolute(value) || value.includes("\\")) throw new Error(`Unsafe generated path: ${value}`);
	const normalized = path.posix.normalize(value);
	if (normalized === ".." || normalized.startsWith("../") || normalized.startsWith("/")) throw new Error(`Unsafe generated path: ${value}`);
	return normalized;
}
/** Materialize a confined, self-contained PPTD source plane inside an empty staging directory. */
async function writePptdProjectSource(directory, source) {
	const manifest = safeOutputRelative(source.entryName);
	await writeFile(path.join(directory, ...manifest.split("/")), source.manifest, {
		encoding: "utf8",
		mode: 384
	});
	for (const [relative, content] of source.pages) {
		const safe = safeOutputRelative(relative);
		const target = path.join(directory, ...safe.split("/"));
		await mkdir(path.dirname(target), {
			recursive: true,
			mode: 448
		});
		await writeFile(target, content, {
			encoding: "utf8",
			mode: 384
		});
	}
	for (const asset of source.assets.values()) {
		const safe = safeOutputRelative(asset.path);
		const target = path.join(directory, ...safe.split("/"));
		await mkdir(path.dirname(target), {
			recursive: true,
			mode: 448
		});
		await writeFile(target, asset.bytes, { mode: 384 });
	}
}
//#endregion
//#region src/pptd-repair.ts
/** Deterministic, explicitly requested PPTD source repairs for `check --level auto`. */
const ELEMENT_TYPES = new Set([
	"text",
	"shape",
	"line",
	"image",
	"icon",
	"table",
	"chart"
]);
function record(value) {
	return typeof value === "object" && value !== null && !Array.isArray(value) ? value : void 0;
}
function action(actions, file, element, kind, field, message) {
	const elementId = typeof element?.elementId === "string" ? element.elementId : void 0;
	actions.push({
		file,
		...elementId === void 0 ? {} : { elementId },
		kind,
		field,
		message
	});
}
function finiteNumber(value) {
	if (typeof value === "number" && Number.isFinite(value)) return value;
	if (typeof value === "string" && value.trim() !== "" && Number.isFinite(Number(value))) return Number(value);
}
function booleanValue(value) {
	if (typeof value === "boolean") return value;
	if (value === "true") return true;
	if (value === "false") return false;
}
function normalizeNumber(owner, field, actions, file, element, range) {
	if (!(field in owner)) return;
	const before = owner[field];
	const converted = finiteNumber(before);
	if (converted !== void 0 && (range === void 0 || range(converted))) {
		if (before !== converted) {
			owner[field] = converted;
			action(actions, file, element, "convert-type", field, `${field} 已转换为数字。`);
		}
		return;
	}
	Reflect.deleteProperty(owner, field);
	action(actions, file, element, "delete-optional-field", field, `${field} 不是有效的可选数字，已删除。`);
}
function normalizeBoolean(owner, field, actions, file, element) {
	if (!(field in owner)) return;
	const before = owner[field];
	const converted = booleanValue(before);
	if (converted !== void 0) {
		if (before !== converted) {
			owner[field] = converted;
			action(actions, file, element, "convert-type", field, `${field} 已转换为布尔值。`);
		}
		return;
	}
	Reflect.deleteProperty(owner, field);
	action(actions, file, element, "delete-optional-field", field, `${field} 不是有效的可选布尔值，已删除。`);
}
function normalizeTuple(owner, field, size, actions, file, element, required) {
	const before = owner[field];
	if (!Array.isArray(before) || before.length !== size) {
		if (!required && field in owner) {
			Reflect.deleteProperty(owner, field);
			action(actions, file, element, "delete-optional-field", field, `${field} 不是长度为 ${size} 的数字数组，已删除。`);
		}
		return false;
	}
	const converted = before.map(finiteNumber);
	if (converted.some((value) => value === void 0)) {
		if (!required) {
			Reflect.deleteProperty(owner, field);
			action(actions, file, element, "delete-optional-field", field, `${field} 包含无效数字，已删除。`);
		}
		return false;
	}
	if (before.some((value, index) => value !== converted[index])) {
		owner[field] = converted;
		action(actions, file, element, "convert-type", field, `${field} 中的数字字符串已转换。`);
	}
	return true;
}
function normalizeBooleanTuple(owner, field, size, actions, file, element) {
	const before = owner[field];
	if (!(field in owner)) return;
	if (!Array.isArray(before) || before.length !== size) {
		Reflect.deleteProperty(owner, field);
		action(actions, file, element, "delete-optional-field", field, `${field} 不是长度为 ${size} 的布尔数组，已删除。`);
		return;
	}
	const converted = before.map(booleanValue);
	if (converted.some((value) => value === void 0)) {
		Reflect.deleteProperty(owner, field);
		action(actions, file, element, "delete-optional-field", field, `${field} 包含无效布尔值，已删除。`);
		return;
	}
	if (before.some((value, index) => value !== converted[index])) {
		owner[field] = converted;
		action(actions, file, element, "convert-type", field, `${field} 中的布尔字符串已转换。`);
	}
}
function normalizeCrop(owner, actions, file, element) {
	if (!("crop" in owner)) return;
	const crop = record(owner.crop);
	if (crop === void 0) {
		Reflect.deleteProperty(owner, "crop");
		action(actions, file, element, "delete-optional-field", "crop", "crop 不是裁剪对象，已删除。");
		return;
	}
	for (const field of [
		"left",
		"top",
		"right",
		"bottom"
	]) normalizeNumber(crop, field, actions, file, element);
}
function normalizeString(owner, field, actions, file, element) {
	const before = owner[field];
	if (typeof before === "string") return true;
	if (typeof before === "number" || typeof before === "boolean") {
		owner[field] = String(before);
		action(actions, file, element, "convert-type", field, `${field} 已转换为字符串。`);
		return true;
	}
	return false;
}
function normalizeStyle(style, actions, file, element) {
	for (const field of [
		"fontSize",
		"lineHeight",
		"letterSpacing",
		"paragraphSpacing",
		"opacity"
	]) normalizeNumber(style, field, actions, file, element, field === "opacity" ? (value) => value >= 0 && value <= 1 : (value) => value >= 0);
	for (const field of [
		"bold",
		"italic",
		"wrap",
		"strike"
	]) normalizeBoolean(style, field, actions, file, element);
}
function normalizeElement(element, actions, file) {
	if (!normalizeString(element, "elementId", actions, file, element) || String(element.elementId).trim() === "") return false;
	if (typeof element.elementType !== "string" || !ELEMENT_TYPES.has(element.elementType)) return false;
	const allowed = ELEMENT_FIELDS[element.elementType];
	if (allowed !== void 0) for (const field of Object.keys(element)) {
		if (allowed.has(field)) continue;
		Reflect.deleteProperty(element, field);
		action(actions, file, element, "delete-optional-field", field, `未知字段 ${field} 已删除。`);
	}
	if (!normalizeTuple(element, "bounds", 4, actions, file, element, true)) return false;
	const bounds = element.bounds;
	if ((bounds[2] ?? -1) < 0 || (bounds[3] ?? -1) < 0) return false;
	for (const field of ["rotation", "opacity"]) normalizeNumber(element, field, actions, file, element, field === "opacity" ? (value) => value >= 0 && value <= 1 : void 0);
	normalizeBooleanTuple(element, "flip", 2, actions, file, element);
	normalizeCrop(element, actions, file, element);
	const type = element.elementType;
	if (type === "text") {
		const content = record(element.content);
		if (content === void 0 || !normalizeString(content, "text", actions, file, element)) return false;
		normalizeStyle(content, actions, file, element);
	}
	if (type === "shape") {
		if (typeof element.shapeName !== "string") return false;
		if (element.shapeName === "custom" && (!normalizeTuple(element, "viewBox", 2, actions, file, element, true) || typeof element.path !== "string")) return false;
	}
	if (type === "line") {
		if (!normalizeTuple(element, "viewBox", 2, actions, file, element, true) || typeof element.points !== "string") return false;
	}
	if (type === "image" && typeof element.src !== "string") return false;
	if (type === "icon" && typeof element.iconName !== "string") return false;
	if (type === "table") {
		if (!Array.isArray(element.rows) || !Array.isArray(element.columnWidths) || !Array.isArray(element.rowHeights)) return false;
		const normalizeRatios = (field) => {
			const values = element[field].map(finiteNumber);
			if (values.length === 0 || values.some((value) => value === void 0 || value < 0 || value > 1)) return false;
			const numeric = values;
			const sum = numeric.reduce((total, value) => total + value, 0);
			if (sum <= 0) return false;
			const normalized = Math.abs(sum - 1) < 1e-9 ? numeric : numeric.map((value) => value / sum);
			const output = Math.abs(sum - 1) < 1e-9 ? normalized : normalized.map((value, index) => index === normalized.length - 1 ? 1 - normalized.slice(0, -1).reduce((total, item) => total + item, 0) : value);
			if (element[field].some((value, index) => value !== output[index])) {
				element[field] = output;
				action(actions, file, element, "convert-type", field, `${field} 已转换为总和为 1 的数字比例。`);
			}
			return true;
		};
		if (!normalizeRatios("columnWidths") || !normalizeRatios("rowHeights")) return false;
	}
	if (type === "chart") {
		const data = record(element.data);
		if (data === void 0 || !Array.isArray(data.cols) || !Array.isArray(data.rows) || !Array.isArray(element.series)) return false;
	}
	return true;
}
function parseYamlObject(content, file) {
	const object = record(yaml.load(content, {
		schema: yaml.JSON_SCHEMA,
		json: true
	}));
	if (object === void 0) throw new Error(`${file} YAML root must be an object before auto repair`);
	return object;
}
function dumpYaml(value) {
	return yaml.dump(value, {
		schema: yaml.JSON_SCHEMA,
		noRefs: true,
		lineWidth: -1,
		sortKeys: false
	});
}
async function writeTextAtomic(target, content) {
	await mkdir(path.dirname(target), { recursive: true });
	const temporary = path.join(path.dirname(target), `.${path.basename(target)}.${randomUUID()}.tmp`);
	const handle = await open(temporary, "wx", 384);
	try {
		await handle.writeFile(content, "utf8");
		await handle.sync();
	} finally {
		await handle.close();
	}
	try {
		await rename(temporary, target);
	} finally {
		await rm(temporary, { force: true });
	}
}
/** Apply bounded deterministic repairs to the selected source pages and recheck the project. */
async function repairPptdProject(inputPath, selectedPages) {
	const entry = await resolvePptdEntry(inputPath);
	const root = path.dirname(entry);
	const entryName = path.basename(entry);
	const actions = [];
	const changed = /* @__PURE__ */ new Map();
	const manifestContent = await readFile(entry, "utf8");
	const manifest = parseYamlObject(manifestContent, entryName);
	const manifestActionStart = actions.length;
	for (const field of Object.keys(manifest)) {
		if (MANIFEST_FIELDS.has(field)) continue;
		Reflect.deleteProperty(manifest, field);
		action(actions, entryName, void 0, "delete-optional-field", field, `未知字段 ${field} 已删除。`);
	}
	normalizeString(manifest, "title", actions, entryName);
	normalizeTuple(manifest, "size", 2, actions, entryName, void 0, false);
	const textStyles = record(record(manifest.theme)?.textStyles);
	if (textStyles !== void 0) {
		const styles = Object.values(textStyles).map(record).filter((value) => value !== void 0);
		for (const style of styles) normalizeStyle(style, actions, entryName);
	}
	const repairedManifest = dumpYaml(manifest);
	if (actions.length > manifestActionStart && repairedManifest !== manifestContent) changed.set(entry, repairedManifest);
	const pageRefs = Array.isArray(manifest.pages) ? manifest.pages : [];
	for (const [index, rawRef] of pageRefs.entries()) {
		if (selectedPages !== void 0 && !selectedPages.has(index + 1)) continue;
		if (typeof rawRef !== "string" || path.isAbsolute(rawRef) || rawRef.includes("\\") || rawRef.startsWith("../")) continue;
		const target = path.join(root, ...path.posix.normalize(rawRef).split("/"));
		let content;
		try {
			content = await readFile(target, "utf8");
		} catch {
			continue;
		}
		const page = parseYamlObject(content, rawRef);
		const pageActionStart = actions.length;
		for (const field of Object.keys(page)) {
			if (PAGE_FIELDS.has(field)) continue;
			Reflect.deleteProperty(page, field);
			action(actions, rawRef, void 0, "delete-optional-field", field, `未知字段 ${field} 已删除。`);
		}
		if (normalizeString(page, "notes", actions, rawRef) === false && page.notes !== void 0) {
			Reflect.deleteProperty(page, "notes");
			action(actions, rawRef, void 0, "delete-optional-field", "notes", "notes 不是可转换的字符串，已删除。");
		}
		const elements = Array.isArray(page.elements) ? page.elements : [];
		const repairedElements = [];
		for (const [elementIndex, rawElement] of elements.entries()) {
			const element = record(rawElement);
			if (element !== void 0 && normalizeElement(element, actions, rawRef)) repairedElements.push(element);
			else action(actions, rawRef, element, "delete-invalid-element", `elements[${elementIndex}]`, "元素缺少无法安全修复的必填结构，已删除。");
		}
		page.elements = repairedElements;
		const repaired = dumpYaml(page);
		if (actions.length > pageActionStart && repaired !== content) changed.set(target, repaired);
	}
	for (const [target, content] of changed) await writeTextAtomic(target, content);
	const project = await loadPptdProject(entry);
	return {
		changedFiles: [...changed.keys()].map((target) => path.relative(root, target) || entryName),
		actions,
		check: checkPptdProject(project)
	};
}
//#endregion
//#region src/bin.ts
/** Local, bounded command line interface for the local PPTD v2 toolchain. */
const MAX_PPTX_BYTES = 512 * 1024 * 1024;
function usage() {
	return [
		"Usage:",
		"  dsh-pptd convert <deck.pptx> [-o <project-directory>] [--strict] [--force] [--json]",
		"  dsh-pptd check <project-directory|deck.pptd> [-p <pages>] [-s <severity>] [--level keep|auto] [--json]",
		"  dsh-pptd inspect <project-directory|deck.pptd> [-p <pages>] [--json]",
		"  dsh-pptd screenshot <project-directory|deck.pptd> [-p <pages>] [-o <directory>] [--scale <number>] [--force] [--json]",
		"  dsh-pptd render <project-directory|deck.pptd> [-p <pages>] [-o <deck.pptx>] [--strict] [--force] [--json]",
		"  dsh-pptd package <project-directory|deck.pptd> [...render options]",
		"  dsh-pptd export <project-directory|deck.pptd> [...render options]",
		"",
		"Page specs are 1-based and support 3, 1,3,5, and 2-10.",
		"Severity specs support all, error, warning, one issue code, or comma-separated issue codes.",
		"Strict conversion refuses to publish a PPTD project when any source feature is normalized or unsupported.",
		"PPTD and local resources stay confined to the project directory. Network resources are disabled."
	].join("\n");
}
function optionValue(argv, index, name) {
	const value = argv[index + 1];
	if (value === void 0 || value.startsWith("-")) throw new Error(`${name} requires a value`);
	return value;
}
function parseArgs(argv) {
	const positional = [];
	let output;
	let page;
	let severity = "all";
	let level = "keep";
	let scale = 2;
	let json = false;
	let force = false;
	let strict = false;
	let help = false;
	let version = false;
	for (let index = 0; index < argv.length; index += 1) {
		const value = argv[index];
		if (value === void 0) continue;
		if (value === "--json") {
			json = true;
			continue;
		}
		if (value === "--force") {
			force = true;
			continue;
		}
		if (value === "--strict") {
			strict = true;
			continue;
		}
		if (value === "--help" || value === "-h") {
			help = true;
			continue;
		}
		if (value === "--version" || value === "-v") {
			version = true;
			continue;
		}
		if (value === "--output" || value === "-o") {
			output = optionValue(argv, index, value);
			index += 1;
			continue;
		}
		if (value === "--page" || value === "-p") {
			page = optionValue(argv, index, value);
			index += 1;
			continue;
		}
		if (value === "--severity" || value === "-s") {
			const selectors = optionValue(argv, index, value).split(",").map((item) => item.trim());
			if (selectors.some((item) => item === "" || !/^[a-z][a-z0-9_-]*$/iu.test(item))) throw new Error("--severity must contain all, error, warning, or comma-separated issue codes");
			severity = selectors.join(",");
			index += 1;
			continue;
		}
		if (value === "--level") {
			const raw = optionValue(argv, index, value);
			if (raw !== "keep" && raw !== "auto") throw new Error("--level must be keep or auto");
			level = raw;
			index += 1;
			continue;
		}
		if (value === "--scale") {
			const raw = optionValue(argv, index, value);
			scale = Number(raw);
			if (!Number.isFinite(scale) || scale <= 0 || scale > 8) throw new Error("--scale must be greater than 0 and at most 8");
			index += 1;
			continue;
		}
		if (value.startsWith("-")) throw new Error(`Unknown option: ${value}`);
		positional.push(value);
	}
	if (positional.length > 2) throw new Error(`Unexpected argument: ${positional[2]}`);
	return {
		...positional[0] === void 0 ? {} : { command: positional[0] },
		...positional[1] === void 0 ? {} : { input: positional[1] },
		...output === void 0 ? {} : { output },
		...page === void 0 ? {} : { page },
		severity,
		level,
		scale,
		json,
		force,
		strict,
		help,
		version
	};
}
async function packageVersion() {
	const content = await readFile(new URL("../package.json", import.meta.url), "utf8");
	const parsed = JSON.parse(content);
	if (typeof parsed.version !== "string") throw new Error("package.json does not declare a version");
	return parsed.version;
}
function writeLine(stream, value) {
	stream.write(`${value}\n`);
}
function printValue(io, value, _json) {
	if (typeof value === "string") writeLine(io.stdout, value);
	else writeLine(io.stdout, JSON.stringify(value, null, 2));
}
function pageNumbers(spec, count) {
	if (count < 1) return [];
	if (spec === void 0) return Array.from({ length: count }, (_value, index) => index + 1);
	const selected = /* @__PURE__ */ new Set();
	for (const rawPart of spec.split(",")) {
		const part = rawPart.trim();
		if (/^\d+$/u.test(part)) {
			selected.add(Number(part));
			continue;
		}
		const range = /^(\d+)-(\d+)$/u.exec(part);
		if (range === null) throw new Error(`Invalid page spec: ${part}`);
		const start = Number(range[1]);
		const end = Number(range[2]);
		if (start > end) throw new Error(`Invalid descending page range: ${part}`);
		for (let value = start; value <= end; value += 1) selected.add(value);
	}
	const values = [...selected].sort((left, right) => left - right);
	const invalid = values.find((value) => !Number.isInteger(value) || value < 1 || value > count);
	if (invalid !== void 0) throw new Error(`Page ${invalid} is outside the available range 1-${count}`);
	if (values.length === 0) throw new Error("Page spec selects no pages");
	return values;
}
function selectedProject(project, pages) {
	if (pages.length === project.pages.length && pages.every((page, index) => page === index + 1)) return project;
	const selected = new Set(pages);
	return {
		...project,
		pages: pages.flatMap((page) => {
			const selectedPage = project.pages[page - 1];
			return selectedPage === void 0 ? [] : [selectedPage];
		}),
		parseIssues: project.parseIssues.filter((issue) => issue.page === void 0 || selected.has(issue.page))
	};
}
function issueMatchesSeverity(issue, severity) {
	const selectors = severity.split(",").map((value) => value.trim().toLowerCase());
	if (selectors.includes("all")) return true;
	return selectors.some((selector) => selector === issue.severity || selector === issue.code.toLowerCase());
}
function checkedForPages(project, pages, severity) {
	const subset = selectedProject(project, pages);
	const raw = checkPptdProject(subset);
	const issues = raw.issues.map((issue) => {
		if (issue.page === void 0 || subset === project) return issue;
		const originalPage = pages[issue.page - 1];
		return originalPage === void 0 ? issue : {
			...issue,
			page: originalPage
		};
	});
	const displayedIssues = issues.filter((issue) => issueMatchesSeverity(issue, severity));
	return {
		...raw,
		issues,
		displayedIssues
	};
}
function humanCheck(checked) {
	const lines = (checked.displayedIssues ?? checked.issues).map((issue) => [
		"[fixed: false]",
		`[${issue.severity === "error" ? "Needs revision" : "Suggestion"}:${issue.code}]`,
		issue.file ?? "",
		issue.page === void 0 ? "" : `page=${issue.page}`,
		issue.elementId === void 0 ? "" : `id="${issue.elementId}"`,
		issue.message
	].filter(Boolean).join(" "));
	lines.push(`${checked.status.toUpperCase()}: ${checked.errorCount} error(s), ${checked.warningCount} warning(s), ${checked.pageCount} page(s)`);
	return lines.join("\n");
}
function humanRepairs(actions) {
	return actions.map((item) => [
		"[fixed: true]",
		`[Repair:${item.kind}]`,
		item.file,
		item.elementId === void 0 ? "" : `id="${item.elementId}"`,
		item.message
	].filter(Boolean).join(" ")).join("\n");
}
function inspectProject(project, pages, checked) {
	const elementTypes = {};
	const resources = {
		assets: project.source.assets.size,
		bytes: 0
	};
	for (const asset of project.source.assets.values()) resources.bytes += asset.bytes.byteLength;
	for (const pageNumber of pages) {
		const page = project.pages[pageNumber - 1];
		if (page === void 0) continue;
		for (const element of page.elements) {
			const type = typeof element.elementType === "string" ? element.elementType : "unknown";
			elementTypes[type] = (elementTypes[type] ?? 0) + 1;
		}
	}
	return {
		title: project.title,
		size: [project.width, project.height],
		selectedPages: pages,
		elementTypes,
		resources,
		compatibility: checked.compatibility,
		check: checked
	};
}
function defaultConvertOutput(input) {
	const resolved = path.resolve(input);
	return path.join(path.dirname(resolved), path.basename(resolved, path.extname(resolved)));
}
async function projectRoot(input) {
	const entry = await resolvePptdEntry(input);
	return {
		entry,
		root: path.dirname(entry)
	};
}
async function defaultScreenshotOutput(input) {
	const { root } = await projectRoot(input);
	return path.join(path.dirname(root), `${path.basename(root)}-screenshots`);
}
async function defaultRenderOutput(input) {
	const { entry, root } = await projectRoot(input);
	return (await lstat(path.resolve(input))).isDirectory() ? path.join(path.dirname(root), `${path.basename(root)}.pptx`) : path.join(root, `${path.basename(entry, ".pptd")}.pptx`);
}
async function commandConvert(args, io) {
	const input = path.resolve(args.input);
	const metadata = await lstat(input);
	if (!metadata.isFile() || path.extname(input).toLowerCase() !== ".pptx") throw new Error("convert input must be a .pptx file");
	if (metadata.size > MAX_PPTX_BYTES) throw new Error(`PPTX exceeds the ${MAX_PPTX_BYTES} byte limit`);
	const converted = await convertPptxToPptd(await readFile(input), path.basename(input));
	const normalizedCount = converted.diagnostics.filter((item) => item.level === "normalized").length;
	const unsupportedCount = converted.diagnostics.filter((item) => item.level === "unsupported").length;
    const placeholderCount = converted.diagnostics.filter(item => item.level === 'placeholder').length;
	if (args.strict && converted.diagnostics.length > 0) throw new Error(`strict conversion requires lossless coverage; received ${normalizedCount} normalized, ${placeholderCount} placeholder and ${unsupportedCount} unsupported diagnostic(s)`);
	const output = path.resolve(args.output ?? defaultConvertOutput(input));
	await publishPptdDirectory(output, args.force, (stage) => writePptdProjectSource(stage, converted.source));
	printValue(io, {
		status: converted.diagnostics.length > 0 ? "warning" : "pass",
		input,
		output,
		slideCount: converted.slideCount,
		sourceNodeCount: converted.sourceNodeCount,
		outputElementCount: converted.outputElementCount,
		extractedAssetCount: converted.extractedAssetCount,
		normalizedCount,
		unsupportedCount,
        placeholderCount,
		diagnostics: converted.diagnostics
	}, args.json);
	return 0;
}
async function commandPptd(args, io) {
	if (args.level === "auto" && args.command !== "check") throw new Error("--level auto is supported only by check");
	let project = await loadPptdProject(args.input);
	const pages = pageNumbers(args.page, project.pages.length);
	const repaired = args.level === "auto" ? await repairPptdProject(args.input, new Set(pages)) : void 0;
	if (repaired !== void 0) project = await loadPptdProject(args.input);
	const checked = checkedForPages(project, pages, args.severity);
	if (args.command === "check") {
		printValue(io, args.json ? {
			...checked,
			repair: repaired === void 0 ? void 0 : {
				changedFiles: repaired.changedFiles,
				actions: repaired.actions
			}
		} : [repaired === void 0 ? "" : humanRepairs(repaired.actions), humanCheck(checked)].filter(Boolean).join("\n"), args.json);
		return checked.status === "fail" ? 1 : 0;
	}
	if (args.command === "inspect") {
		printValue(io, inspectProject(project, pages, checked), args.json);
		return checked.status === "fail" ? 1 : 0;
	}
	if (checked.status === "fail" || args.strict && checked.status !== "pass") {
		printValue(io, { ...checked, exported: false, status: "needs_revision" }, args.json,
			"No output produced; validation needs revision.\n" + humanCheck(checked));
		return 1;
	}
	if (args.command === "screenshot") {
		const output = path.resolve(args.output ?? await defaultScreenshotOutput(args.input));
		await publishPptdDirectory(output, args.force, async (stage) => {
			await mkdir(path.join(stage, "pages"), {
				recursive: true,
				mode: 448
			});
			const manifest = [];
			for (const page of pages) {
				const bytes = await renderPptdPagePng(project, page - 1, args.scale);
				const file = `pages/page-${page}.png`;
				await writeFile(path.join(stage, ...file.split("/")), bytes, { mode: 384 });
				manifest.push({
					page,
					file,
					bytes: bytes.byteLength
				});
			}
			const index = `${JSON.stringify({
				title: project.title,
				scale: args.scale,
				pages: manifest
			}, null, 2)}\n`;
			await writeFile(path.join(stage, "index.json"), index, { mode: 384 });
		});
		printValue(io, {
			status: checked.status,
			input: path.resolve(args.input),
			output,
			pages,
			scale: args.scale,
			warnings: checked.warningCount
		}, args.json);
		return 0;
	}
	const output = path.resolve(args.output ?? await defaultRenderOutput(args.input));
	const rendered = await renderPptdProject(selectedProject(project, pages));
	await publishPptdFile(output, rendered.bytes, args.force);
	printValue(io, {
		status: checked.status,
		input: path.resolve(args.input),
		output,
		pages,
		pageCount: pages.length,
		nativeObjectCount: rendered.nativeObjectCount,
		sizeBytes: rendered.bytes.byteLength,
		digest: checked.digest,
		warnings: checked.warningCount,
		compatibility: checked.compatibility
	}, args.json);
	return 0;
}
/** Execute the CLI without terminating the host process. */
async function runCli(argv, io = process) {
	try {
		const args = parseArgs(argv);
		if (args.version) {
			writeLine(io.stdout, await packageVersion());
			return 0;
		}
		if (args.help || args.command === void 0) {
			writeLine(io.stdout, usage());
			return args.help ? 0 : 2;
		}
		if (args.input === void 0) throw new Error(`${args.command} requires an input path\n${usage()}`);
		const command = args.command === "package" || args.command === "export" ? "render" : args.command;
		const normalized = {
			...args,
			command,
			input: args.input
		};
		if (command === "convert") return await commandConvert(normalized, io);
		if (![
			"check",
			"inspect",
			"screenshot",
			"render"
		].includes(command)) throw new Error(`Unknown command: ${args.command}\n${usage()}`);
		return await commandPptd(normalized, io);
	} catch (error) {
		writeLine(io.stderr, error instanceof Error ? error.message : String(error));
		return 1;
	}
}
async function main() {
	process.exitCode = await runCli(process.argv.slice(2), process);
}
const invoked = process.argv[1];
// npm installs .bin symlinks on POSIX; compare canonical paths so that entry runs too.
if (invoked !== void 0 && pathToFileURL(await realpath(path.resolve(invoked)).catch(() => path.resolve(invoked))).href === import.meta.url) await main();
//#endregion
export { runCli };
