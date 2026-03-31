/**
 * Unit tests for parseImportedStyleXml — extracts rPr/pPr from raw style XML.
 */
import { describe, test, expect } from 'bun:test';
import { parseImportedStyleXml } from './parseImportedStyleXml';

describe('parseImportedStyleXml', () => {
  describe('color parsing', () => {
    test('auto color is parsed as auto flag, not rgb', () => {
      const xml = '<w:style><w:rPr><w:color w:val="auto"/></w:rPr></w:style>';
      const result = parseImportedStyleXml(xml);
      expect(result.rPr?.color?.auto).toBe(true);
      expect(result.rPr?.color?.rgb).toBeUndefined();
    });

    test('explicit RGB color is parsed correctly', () => {
      const xml = '<w:style><w:rPr><w:color w:val="FF0000"/></w:rPr></w:style>';
      const result = parseImportedStyleXml(xml);
      expect(result.rPr?.color?.rgb).toBe('FF0000');
      expect(result.rPr?.color?.auto).toBeUndefined();
    });

    test('theme color is parsed with themeColor and themeShade', () => {
      const xml =
        '<w:style><w:rPr><w:color w:val="365F91" w:themeColor="accent1" w:themeShade="BF"/></w:rPr></w:style>';
      const result = parseImportedStyleXml(xml);
      expect(result.rPr?.color?.rgb).toBe('365F91');
      expect(result.rPr?.color?.themeColor).toBe('accent1');
      expect(result.rPr?.color?.themeShade).toBe('BF');
    });

    test('theme color with themeTint is parsed', () => {
      const xml =
        '<w:style><w:rPr><w:color w:val="4472C4" w:themeColor="accent1" w:themeTint="99"/></w:rPr></w:style>';
      const result = parseImportedStyleXml(xml);
      expect(result.rPr?.color?.themeColor).toBe('accent1');
      expect(result.rPr?.color?.themeTint).toBe('99');
    });
  });

  describe('bold parsing', () => {
    test('w:b/ is parsed as bold true', () => {
      const xml = '<w:style><w:rPr><w:b/></w:rPr></w:style>';
      const result = parseImportedStyleXml(xml);
      expect(result.rPr?.bold).toBe(true);
    });

    test('w:b val=false is parsed as bold false (not true)', () => {
      const xml = '<w:style><w:rPr><w:b w:val="false"/></w:rPr></w:style>';
      const result = parseImportedStyleXml(xml);
      expect(result.rPr?.bold).toBe(false);
    });

    test('w:b val=0 is parsed as bold false', () => {
      const xml = '<w:style><w:rPr><w:b w:val="0"/></w:rPr></w:style>';
      const result = parseImportedStyleXml(xml);
      expect(result.rPr?.bold).toBe(false);
    });

    test('no w:b element means bold is undefined', () => {
      const xml = '<w:style><w:rPr><w:sz w:val="24"/></w:rPr></w:style>';
      const result = parseImportedStyleXml(xml);
      expect(result.rPr?.bold).toBeUndefined();
    });
  });

  describe('italic parsing', () => {
    test('w:i/ is parsed as italic true', () => {
      const xml = '<w:style><w:rPr><w:i/></w:rPr></w:style>';
      const result = parseImportedStyleXml(xml);
      expect(result.rPr?.italic).toBe(true);
    });

    test('w:i val=false is parsed as italic false', () => {
      const xml = '<w:style><w:rPr><w:i w:val="false"/></w:rPr></w:style>';
      const result = parseImportedStyleXml(xml);
      expect(result.rPr?.italic).toBe(false);
    });
  });

  describe('font parsing', () => {
    test('theme font references are parsed', () => {
      const xml =
        '<w:style><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/></w:rPr></w:style>';
      const result = parseImportedStyleXml(xml);
      expect(result.rPr?.fontFamily?.asciiTheme).toBe('majorHAnsi');
      expect(result.rPr?.fontFamily?.hAnsiTheme).toBe('majorHAnsi');
    });

    test('explicit font names are parsed', () => {
      const xml =
        '<w:style><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/></w:rPr></w:style>';
      const result = parseImportedStyleXml(xml);
      expect(result.rPr?.fontFamily?.ascii).toBe('Calibri');
      expect(result.rPr?.fontFamily?.hAnsi).toBe('Calibri');
    });
  });

  describe('real heading style XML', () => {
    test('Heading1 with auto color and theme font', () => {
      const xml = `<w:style w:type="paragraph" w:styleId="Heading1">
        <w:name w:val="heading 1"/>
        <w:basedOn w:val="Normal"/>
        <w:pPr><w:spacing w:before="240" w:after="0"/></w:pPr>
        <w:rPr>
          <w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
          <w:b/>
          <w:color w:val="auto"/>
          <w:sz w:val="32"/>
        </w:rPr>
      </w:style>`;
      const result = parseImportedStyleXml(xml);
      expect(result.rPr?.bold).toBe(true);
      expect(result.rPr?.color?.auto).toBe(true);
      expect(result.rPr?.fontFamily?.asciiTheme).toBe('majorHAnsi');
      expect(result.rPr?.fontSize).toBe(32);
      expect(result.pPr?.spacing?.before).toBe(240);
    });

    test('Heading2 without bold (inherits from parent, should not be bold)', () => {
      const xml = `<w:style w:type="paragraph" w:styleId="Heading2">
        <w:name w:val="heading 2"/>
        <w:basedOn w:val="Heading1"/>
        <w:rPr>
          <w:b w:val="false"/>
          <w:color w:val="auto"/>
          <w:sz w:val="26"/>
        </w:rPr>
      </w:style>`;
      const result = parseImportedStyleXml(xml);
      expect(result.rPr?.bold).toBe(false);
      expect(result.rPr?.fontSize).toBe(26);
    });

    test('Heading5 with theme color', () => {
      const xml = `<w:style w:type="paragraph" w:styleId="Heading5">
        <w:rPr>
          <w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/>
          <w:color w:val="365F91" w:themeColor="accent1" w:themeShade="BF"/>
          <w:sz w:val="22"/>
        </w:rPr>
      </w:style>`;
      const result = parseImportedStyleXml(xml);
      expect(result.rPr?.color?.rgb).toBe('365F91');
      expect(result.rPr?.color?.themeColor).toBe('accent1');
      expect(result.rPr?.color?.themeShade).toBe('BF');
    });
  });
});
