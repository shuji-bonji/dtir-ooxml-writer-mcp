#!/usr/bin/env node
/**
 * dtir-ooxml-writer-mcp MCP server
 *
 * tool: dtir_to_docx — 翻訳済み DTIR ＋ 元 .docx(base64) から訳 .docx(base64) を生成。
 * 元ファイルを基板に id でパッチするため、書式・画像・sectPr は保持される。
 */
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import { dtirToDocx } from './writer.js';
import type { IRDocument } from './dtir.js';

const server = new McpServer({ name: 'dtir-ooxml-writer-mcp', version: '0.0.1' });

server.tool(
  'dtir_to_docx',
  '翻訳済み DTIR と元 .docx から訳 .docx を生成する。元ファイルを基板に anchor の id で ' +
    'パッチ（collapse 既定）。translatable=false と IR に乗らない要素（書式/画像/sectPr）は不可触。',
  {
    dtirJson: z.string().describe('翻訳済み DTIR(IRDocument) の JSON 文字列'),
    originalDocxBase64: z.string().describe('元 .docx の base64'),
    onMissingTranslation: z
      .enum(['keep', 'error'])
      .optional()
      .describe('未翻訳セグメントの扱い（既定 keep=原文維持）'),
  },
  async (args) => {
    try {
      const dtir = JSON.parse(args.dtirJson) as IRDocument;
      const orig = Buffer.from(args.originalDocxBase64, 'base64');
      const out = await dtirToDocx(dtir, orig, {
        onMissingTranslation: args.onMissingTranslation,
      });
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              fileName: dtir.source.fileName,
              byteSize: out.length,
              docxBase64: out.toString('base64'),
            }),
          },
        ],
      };
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { isError: true, content: [{ type: 'text', text: `dtir_to_docx failed: ${msg}` }] };
    }
  },
);

const transport = new StdioServerTransport();
await server.connect(transport);
console.error('dtir-ooxml-writer-mcp MCP server running on stdio');
