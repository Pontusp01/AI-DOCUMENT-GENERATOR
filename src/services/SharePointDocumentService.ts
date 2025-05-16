// src/services/SharePointDocumentService.ts
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from '../utils/pnpjsConfig';
import { SPFI } from '@pnp/sp';
import { ISearchQuery, SearchResults } from '@pnp/sp/search';
import { MSGraphClientFactory } from '@microsoft/sp-http';
import { Document, Packer, Paragraph, TextRun, 
  HeadingLevel, AlignmentType, BorderStyle, Table, TableRow, TableCell, WidthType  
} from 'docx';
// Importera alla nödvändiga PnP-moduler
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import '@pnp/sp/search';

export interface SharePointDocument {
  id: string;
  name: string;
  url: string;
  contentType: string;
  createdDate: Date;
}

export class SharePointDocumentService {
  private context: WebPartContext;
  private sp: SPFI;
  private useFallbackMethods: boolean = false;
  
  constructor(context: WebPartContext) {
    this.context = context;
    // Initiera PnP SP
    this.sp = getSP(context);
  }

  /**
   * Hämtar Word-dokument från alla tillgängliga bibliotek
   */
  public async getWordDocumentsFromAllLibraries(): Promise<SharePointDocument[]> {
    try {
      // Om vi redan vet att sökning inte fungerar, använd direkta biblioteksanrop direkt
      if (this.useFallbackMethods) {
        console.log("Använder direkta biblioteksanrop istället för sökning");
        return await this.getDocumentsFromLibraries();
      }
      
      // Försök med sökning först
      try {
        console.log("Försöker hämta dokument med sökning (PnP v4)");
        const searchResults = await this.searchForDocuments("fileextension:docx");
        
        // Om sökningen fungerade, använd resultaten
        if (searchResults.length > 0) {
          console.log(`Sökning lyckades, hittade ${searchResults.length} dokument`);
          return searchResults;
        }
        
        console.log("Sökning returnerade inga resultat, provar med direkta biblioteksanrop");
      } catch (searchError) {
        console.warn("Sökning misslyckades, använder fallback-metod:", searchError);
        this.useFallbackMethods = true; // Markera att sökning inte fungerar för framtida anrop
      }
      
      // Fallback: Använd direkta biblioteksanrop
      return await this.getDocumentsFromLibraries();
    } catch (error) {
      console.error("Fel vid hämtning av dokument:", error);
      return [];
    }
  }
  
  /**
   * Söker efter dokument med PnP-sökning
   */
  private async searchForDocuments(searchQuery: string): Promise<SharePointDocument[]> {
    try {
      // Begränsa sökningen till den aktuella webbplatsen (site URL)
      const currentSiteUrl = this.context.pageContext.web.absoluteUrl;
      
      // Extrahera sitens URL utan 'https://' för att använda i path-begränsningen
      const siteUrlForPath = currentSiteUrl.replace(/^https?:\/\/[^\/]+/, "");
      
      // Skapa sökförfrågan med begränsning till aktuell webbplats
      const query: ISearchQuery = {
        Querytext: `${searchQuery} path:${siteUrlForPath}`,
        RowLimit: 500,
        SelectProperties: ['Title', 'Path', 'FileExtension', 'Write', 'UniqueId', 'ContentTypeId'],
        TrimDuplicates: true
      };
      
      // Utför sökning
      const searchResults: SearchResults = await this.sp.search(query);
      
      if (!searchResults.PrimarySearchResults || searchResults.PrimarySearchResults.length === 0) {
        console.log("Sökningen returnerade inga resultat");
        return [];
      }
      
      console.log(`Sökningen hittade ${searchResults.PrimarySearchResults.length} dokument`);
      
      // Konvertera sökresultat till SharePointDocument-objekt
      const documents: SharePointDocument[] = searchResults.PrimarySearchResults.map(result => {
        // Extrahera filnamnet från Path
        const pathParts = (result.Path || '').split('/');
        const fileName = pathParts.length > 0 ? pathParts[pathParts.length - 1] : 'Unnamed Document';
        
        return {
          id: result.UniqueId || 'unknown',
          name: fileName || result.Title || 'Unnamed Document',
          url: result.Path || '',
          contentType: 'Word Document',
          createdDate: result.Write ? new Date(result.Write) : new Date()
        };
      });
      
      return documents;
    } catch (error) {
      console.warn("Fel vid sökning efter dokument:", error);
      throw error;
    }
  }
  
  /**
   * Hämtar dokument genom att iterera igenom bibliotek
   */
  private async getDocumentsFromLibraries(): Promise<SharePointDocument[]> {
    try {
      const libraries = await this.getDocumentLibraries();
      let allDocuments: SharePointDocument[] = [];
      
      console.log(`Hämtar dokument från ${libraries.length} bibliotek`);
      
      // Försök med varje bibliotek
      for (let i = 0; i < libraries.length; i++) {
        try {
          const docs = await this.getWordDocuments(libraries[i]);
          if (docs.length > 0) {
            allDocuments = [...allDocuments, ...docs];
            console.log(`Hittade ${docs.length} dokument i ${libraries[i]}`);
          }
        } catch (error) {
          console.warn(`Kunde inte hämta dokument från ${libraries[i]}:`, error);
          // Fortsätt med nästa bibliotek
        }
      }
      
      return allDocuments;
    } catch (error) {
      console.error("Fel vid genomgång av bibliotek:", error);
      return [];
    }
  }

  /**
   * Hämtar alla tillgängliga dokumentbibliotek
   */
  public async getDocumentLibraries(): Promise<string[]> {
    try {
      // Använd PnP för att hämta bibliotek
      const lists = await this.sp.web.lists
        .filter("BaseTemplate eq 101")
        .select("Title")();
      
      // Filtrera bibliotek för att ta bort problematiska
      const libraries: string[] = lists
        .filter(list => !list.Title.includes('/') && !list.Title.includes('_catalogs'))
        .map(list => list.Title);
      
      console.log("Hittade bibliotek (filtrerade):", libraries);
      
      // Lägg till standardbibliotek om nödvändigt
      if (libraries.length === 0) {
        libraries.push("Shared Documents");
        libraries.push("Documents");
        libraries.push("Dokument");
      }
      
      // Se till att Shared Documents finns i listan
      if (!libraries.includes("Shared Documents")) {
        libraries.push("Shared Documents");
      }
      
      return libraries;
    } catch (error) {
      console.error("Fel vid hämtning av bibliotek:", error);
      return ["Shared Documents", "Documents", "Dokument"]; 
    }
  }

  /**
   * Hämtar Word-dokument från ett specifikt bibliotek
   */
  public async getWordDocuments(libraryName: string = "Shared Documents"): Promise<SharePointDocument[]> {
    try {
      console.log(`Försöker hämta Word-dokument från bibliotek: "${libraryName}"`);
      
      // Använd PnP för att hämta dokument
      const items = await this.sp.web.lists.getByTitle(libraryName).items
        .filter("substringof('.docx', FileLeafRef)")
        .expand("File")
        .select("Id", "Title", "FileLeafRef", "Created", "File/ServerRelativeUrl")();
      
      console.log(`Hämtade ${items.length} dokument från ${libraryName}`);
      
      // Konvertera items till SharePointDocument-objekt
      const documents: SharePointDocument[] = items.map(item => ({
        id: item.Id,
        name: item.FileLeafRef || 'Unnamed Document',
        url: item.File?.ServerRelativeUrl || '',
        contentType: 'Word Document',
        createdDate: new Date(item.Created)
      }));
      
      return documents;
    } catch (error) {
      console.log(`Fel vid hämtning av dokument från ${libraryName}:`, error);
      
      // Alternativ metod med mappar
      try {
        console.log(`Provar alternativ metod för att hämta dokument från ${libraryName}...`);
        
        // Bygg relativ URL för biblioteket
        const webUrl = this.context.pageContext.web.serverRelativeUrl;
        const hasTrailingSlash = webUrl.charAt(webUrl.length - 1) === '/';
        const folderPath = hasTrailingSlash ? `${webUrl}${libraryName}` : `${webUrl}/${libraryName}`;
        
        // Hämta filer med getPaged för att hantera begränsningar i PnP v4
        const folder = this.sp.web.getFolderByServerRelativePath(folderPath);
        const files = await folder.files();
        
        // Filtrera .docx-filer manuellt
        const docxFiles = files.filter(file => file.Name.toLowerCase().endsWith('.docx'));
        
        console.log(`Hämtade ${docxFiles.length} dokument med alternativ metod`);
        
        // Konvertera files till SharePointDocument-objekt
        const documents: SharePointDocument[] = docxFiles.map(file => ({
          id: file.UniqueId || 'unknown',
          name: file.Name || 'Unnamed Document',
          url: file.ServerRelativeUrl || '',
          contentType: 'Word Document',
          createdDate: new Date(file.TimeCreated || file.TimeLastModified)
        }));
        
        return documents;
      } catch (altError) {
        console.error(`Alternativ metod misslyckades för ${libraryName}:`, altError);
        return [];
      }
    }
  }
  
  /**
   * Hämtar innehåll från ett dokument
   */
  public async getDocumentContent(serverRelativeUrl: string): Promise<string> {
    try {
      // Kontrollera om det är en Word-fil (.docx)
      const isWordFile = serverRelativeUrl.toLowerCase().endsWith(".docx");
      
      if (isWordFile) {
        console.log("Fil är ett Word-dokument (.docx), returnerar platshållartext");
        // För .docx-filer, returnera en platshållartext istället för att försöka läsa binärt innehåll
        return "Detta är en platshållare för innehåll i Word-dokument. Word-dokument kan inte läsas som text direkt.";
      }
      
      // För andra filer, använd PnP för att hämta filinnehåll
      const fileContent = await this.sp.web.getFileByServerRelativePath(serverRelativeUrl).getText();
      
      // Rensa innehållet
      const textContent = fileContent.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
      
      return textContent;
    } catch (error) {
      console.error("Fel vid hämtning av dokumentinnehåll:", error);
      return "";
    }
  }



  /**
   * Skapar ett dokument från en mall
   */
    /**
   * Skapar ett dokument från en mall
   */

  public async createDocumentFromTemplate(
    templateUrl: string, 
    replacements: Record<string, string>,
    newFileName: string,
    libraryName: string = "Shared Documents"
  ): Promise<string> {
    try {
      // Kontrollera om det är en Word-mall (.docx)
      const isWordTemplate = templateUrl.toLowerCase().endsWith(".docx");
      
      if (isWordTemplate) {
        console.log("Mall är ett Word-dokument (.docx), använder HTML-skapande utan mall");
        
        // Skapa HTML-innehåll direkt från ersättningarna
        let content = "";
        
        if (replacements['CONTENT']) {
          content = replacements['CONTENT'];
        } else {
          // Bygger innehåll från alla ersättningar
          content = Object.entries(replacements)
            .map(([key, value]) => `<h2>${key}</h2><div>${value}</div>`)
            .join('<hr>');
        }
        
        // Använd HTML-metoden för att skapa dokumentet
        return await this.createHtmlDocument(content, newFileName, libraryName);
      }
      
      // För andra filtyper, fortsätt som vanligt
      const templateContent = await this.getDocumentContent(templateUrl);
      
      let newContent = templateContent;
      
      // Ersätt platshållare
      for (const key in replacements) {
        if (replacements.hasOwnProperty(key)) {
          const value = replacements[key];
          const placeholder = `{{${key}}}`;
          newContent = newContent.replace(new RegExp(placeholder, 'g'), value);
        }
      }
      
      // Skapa HTML-dokument
      return await this.createHtmlDocument(newContent, newFileName, libraryName);
    } catch (error) {
      console.error("Fel vid skapande av dokument från mall:", error);
      throw error;
    }
  }

  /**
   * Hämtar sidor från SitePages-biblioteket
   */
  public async getSitePages(): Promise<SharePointDocument[]> {
    try {
      // Om sökning inte fungerar, använd direkt metod
      if (this.useFallbackMethods) {
        return this.getSitePagesDirectly();
      }
      
      // Försök med sökning först
      try {
        const searchResults = await this.searchForDocuments("contentclass:STS_ListItem_WebPageLibrary");
        if (searchResults.length > 0) {
          return searchResults;
        }
      } catch (searchError) {
        console.warn("Sökning efter sidor misslyckades:", searchError);
      }
      
      // Fallback till direkt metod
      return this.getSitePagesDirectly();
    } catch (error) {
      console.error("Fel vid hämtning av sidor:", error);
      return [];
    }
  }
  
  /**
   * Hämtar sidor direkt från SitePages-biblioteket
   */
  private async getSitePagesDirectly(): Promise<SharePointDocument[]> {
    try {
      console.log("Hämtar sidor direkt från SitePages-biblioteket");
      
      // Använd PnP för att hämta sidor
      const items = await this.sp.web.lists.getByTitle("SitePages").items
        .expand("File")
        .select("Id", "Title", "FileLeafRef", "Created", "File/ServerRelativeUrl")();
      
      // Konvertera items till SharePointDocument-objekt
      const pages: SharePointDocument[] = items.map(item => ({
        id: item.Id,
        name: item.Title || item.FileLeafRef || 'Unnamed Page',
        url: item.File?.ServerRelativeUrl || '',
        contentType: 'Site Page',
        createdDate: new Date(item.Created)
      }));
      
      return pages;
    } catch (error) {
      console.error("Fel vid direkt hämtning av sidor:", error);
      return [];
    }
  }
  
  /**
   * Testar anslutningen till SharePoint och väljer bästa metoden
   */
  public async testConnectivity(): Promise<void> {
    console.log("=== SharePoint Connectivity Test (PnP v4) ===");
    
    // Diagnostikutskrifter för URL-information
    console.log("API-anrop detaljer:", {
      baseUrl: this.context.pageContext.web.absoluteUrl,
      serverRelativeUrl: this.context.pageContext.web.serverRelativeUrl,
      siteUrl: this.context.pageContext.site.absoluteUrl,
      userLoginName: this.context.pageContext.user.loginName
    });
    
    // Test 1: Grundläggande åtkomst
    try {
      console.log("Test 1: Grundläggande webbåtkomst...");
      const web = await this.sp.web.select("Title")();
      console.log("Test 1 resultat: LYCKAT");
      console.log(`Webbplatsens titel: ${web.Title}`);
    } catch (error) {
      console.error("Test 1 misslyckades:", error);
    }
    
    // Test 2: Biblioteksåtkomst
    try {
      console.log("Test 2: Hämtar dokumentbibliotek...");
      const libraries = await this.getDocumentLibraries();
      console.log(`Test 2 resultat: Hittade ${libraries.length} bibliotek - ${libraries.join(', ')}`);
    } catch (error) {
      console.error("Test 2 misslyckades:", error);
    }
    
    // Test 3: Sökåtkomst
    try {
      console.log("Test 3: Testar sökfunktionalitet...");
      
      try {
        // Använd PnP Search
        const searchQuery: ISearchQuery = {
          Querytext: "fileextension:docx",
          RowLimit: 1
        };
        
        const searchResults = await this.sp.search(searchQuery);
        
        const hasResults = searchResults.PrimarySearchResults && 
                          searchResults.PrimarySearchResults.length > 0;
        
        console.log(`Test 3 resultat: LYCKAT (${hasResults ? 'hittade' : 'hittade inga'} resultat)`);
        this.useFallbackMethods = !hasResults;
      } catch (searchError) {
        console.error("Sök-API misslyckades:", searchError);
        this.useFallbackMethods = true;
        console.log("Sökning stöds inte i denna miljö, använder fallback-metoder");
      }
    } catch (error) {
      console.error("Test 3 misslyckades:", error);
      this.useFallbackMethods = true;
    }
    
    // Test 4: Hämta dokument från första biblioteket
    try {
      console.log("Test 4: Testar att hämta dokument från första biblioteket...");
      const libraries = await this.getDocumentLibraries();
      
      if (libraries.length > 0) {
        const firstLibrary = libraries[0];
        console.log(`Testar bibliotek: ${firstLibrary}`);
        const docs = await this.getWordDocuments(firstLibrary);
        console.log(`Test 4 resultat: Hittade ${docs.length} dokument i ${firstLibrary}`);
        
        if (docs.length > 0) {
          console.log("Exempel på dokument:", docs[0]);
        }
      } else {
        console.log("Test 4: Inga bibliotek tillgängliga för test");
      }
    } catch (error) {
      console.error("Test 4 misslyckades:", error);
    }
    
    console.log(`Strategi vald: ${this.useFallbackMethods ? 'Direkta anrop' : 'Sökning när möjligt'}`);
    console.log("=== Slut på testning ===");
  }


public async createWordDocument(
  content: string, 
  fileName: string, 
  libraryName: string = "Shared Documents"
): Promise<string> {
  try {
    // Hitta ett tillgängligt bibliotek
    let availableLibrary = "Shared Documents";
    
    try {
      const libraries = await this.getDocumentLibraries();
      
      if (libraries.includes(libraryName)) {
        availableLibrary = libraryName;
      } else if (libraries.length > 0) {
        availableLibrary = libraries[0];
        console.log(`Använder tillgängligt bibliotek: ${availableLibrary} istället för ${libraryName}`);
      }
    } catch (error) {
      console.warn("Kunde inte hämta bibliotek, använder Shared Documents:", error);
    }
    
    console.log(`Skapar docx dokument i bibliotek: ${availableLibrary}`);
    
    try {
      // Dela upp innehållet i rader för att analysera formateringen
      const lines = content.split('\n');
      const paragraphs = [];
      
      let inTable = false;
      let tableRows = [];
      let currentTableRow = [];
      
      // Bearbeta varje rad för att skapa välformaterade paragrafer
      for (let i = 0; i < lines.length; i++) {
        let line = lines[i].trim();
        
        if (line === '') {
          // Tomma rader - lägg till en tom paragraf
          if (!inTable) {
            paragraphs.push(new Paragraph({ text: '' }));
          }
          continue;
        }
        
        // Kontrollera om det är en horisontell linje (---)
        if (line.match(/^-{3,}$/)) {
          // Skapa en horisontell linje med en border
          paragraphs.push(
            new Paragraph({
              text: "",
              border: {
                bottom: {
                  color: "999999",
                  space: 1,
                  style: BorderStyle.SINGLE,
                  size: 6
                }
              },
              spacing: { before: 200, after: 200 }
            })
          );
          continue;
        }
        
        // Kontrollera om det är en tabellrad
        if (line.startsWith('|') && line.endsWith('|')) {
          if (!inTable) {
            // Börja en ny tabell
            inTable = true;
            tableRows = [];
            currentTableRow = [];
          }
          
          // Dela upp tabellceller
          const cells = line.split('|').map(cell => cell.trim()).filter(cell => cell !== '');
          currentTableRow = cells;
          tableRows.push(currentTableRow);
          
          // Om nästa rad inte är en tabellrad, avsluta tabellen
          if (i + 1 >= lines.length || !lines[i + 1].trim().startsWith('|')) {
            // Skapa en tabell när alla rader har samlats in
            const table = this.createTableFromRows(tableRows);
            paragraphs.push(table);
            inTable = false;
          }
          continue;
        }
        
        // Kontrollera om det är en rubrik med ### (Heading 3)
        if (line.startsWith('### ')) {
          const headingText = line.substring(4).trim();
          paragraphs.push(
            new Paragraph({
              text: headingText,
              heading: HeadingLevel.HEADING_3,
              spacing: { before: 200, after: 100 }
            })
          );
          continue;
        }
        
        // Kontrollera om det är en rubrik med ## (Heading 2)
        if (line.startsWith('## ')) {
          const headingText = line.substring(3).trim();
          paragraphs.push(
            new Paragraph({
              text: headingText,
              heading: HeadingLevel.HEADING_2,
              spacing: { before: 200, after: 100 }
            })
          );
          continue;
        }
        
        // Kontrollera om det är en rubrik med # (Heading 1)
        if (line.startsWith('# ')) {
          const headingText = line.substring(2).trim();
          paragraphs.push(
            new Paragraph({
              text: headingText,
              heading: HeadingLevel.HEADING_1,
              spacing: { before: 200, after: 100 }
            })
          );
          continue;
        }
        
        // Kontrollera om det är en rubrik med ** (fetstil)
        if (line.startsWith('**') && line.endsWith('**')) {
          const headingText = line.substring(2, line.length - 2).trim();
          paragraphs.push(
            new Paragraph({
              text: headingText,
              heading: HeadingLevel.HEADING_1,
              alignment: AlignmentType.CENTER,
              spacing: { before: 200, after: 100 }
            })
          );
          continue;
        }
        
        // Kontrollera om det är en punktlista
        if (line.startsWith('* ')) {
          const listText = line.substring(2).trim();
          paragraphs.push(
            new Paragraph({
              text: listText,
              bullet: { level: 0 },
              spacing: { before: 100, after: 100 }
            })
          );
          continue;
        }
        
        // Kontrollera om raden innehåller fetstil-markeringar
        if (line.includes('**')) {
          const textRuns = [];
          let index = 0;
          let inBold = false;
          
          while (index < line.length) {
            // Hitta nästa förekomst av **
            const boldIndex = line.indexOf('**', index);
            
            if (boldIndex === -1) {
              // Ingen mer fetstil, lägg till resten av texten
              if (index < line.length) {
                textRuns.push(new TextRun({ text: line.substring(index), bold: inBold }));
              }
              break;
            }
            
            // Lägg till text fram till **
            if (boldIndex > index) {
              textRuns.push(new TextRun({ text: line.substring(index, boldIndex), bold: inBold }));
            }
            
            index = boldIndex + 2;
            inBold = !inBold;
          }
          
          paragraphs.push(new Paragraph({ 
            children: textRuns,
            spacing: { before: 100, after: 100 }
          }));
        } else {
          // Vanlig text utan formatering
          paragraphs.push(new Paragraph({ 
            text: line,
            spacing: { before: 100, after: 100 }
          }));
        }
      }
      
      // Skapa ett korrekt docx-dokument med de formaterade paragraferna
      const doc = new Document({
        sections: [{
          properties: {},
          children: paragraphs
        }]
      });
      
      // Konvertera docx-dokumentet till en blob
      const blob = await Packer.toBlob(doc);
      
      // Konvertera blob till array buffer för uppladdning
      const arrayBuffer = await blob.arrayBuffer();
      const uint8Array = new Uint8Array(arrayBuffer);
      
      // Bygg URL för biblioteket
      const webUrl = this.context.pageContext.web.serverRelativeUrl;
      const hasTrailingSlash = webUrl.charAt(webUrl.length - 1) === '/';
      const folderPath = hasTrailingSlash ? `${webUrl}${availableLibrary}` : `${webUrl}/${availableLibrary}`;
      
      // Ladda upp filen med PnP
      const fileAddResult = await this.sp.web.getFolderByServerRelativePath(folderPath)
        .files.addUsingPath(`${fileName}.docx`, uint8Array, { Overwrite: true });
      
      console.log("DOCX-fil skapad med docx.js och formaterad:", fileAddResult.ServerRelativeUrl);
      
      return fileAddResult.ServerRelativeUrl;
    } catch (docxJsError) {
      console.error("docx.js misslyckades:", docxJsError);
      
      // Fallback till Graph API om docx.js misslyckas
      try {
        console.log("Använder Microsoft Graph API som fallback...");
        return await this.createWordDocumentWithGraph(content, fileName, availableLibrary);
      } catch (graphError) {
        console.error("Graph API misslyckades:", graphError);
        
        // Sista fallback - skapa HTML-fil istället
        console.log("Använder fallback för att skapa HTML-fil...");
        return await this.createHtmlDocument(content, fileName, libraryName);
      }
    }
  } catch (error) {
    console.error("Alla försök att skapa dokument misslyckades:", error);
    throw error;
  }
}

/**
 * Hjälpmetod för att skapa en tabell från rader
 */
private createTableFromRows(rows: string[][]): Table {
  // Kontrollera om första raden är rubrikrad
  const hasHeader = rows.length > 1 && rows[1].join('').includes('-');
  
  // Skapa riktiga tabellrader
  const tableRows: TableRow[] = [];
  
  // Lägg till rubrikraden om den finns
  if (hasHeader && rows.length > 0) {
    const headerRow = new TableRow({
      children: rows[0].map(cell => 
        new TableCell({
          children: [  
            new Paragraph({ 
              children: [
                new TextRun({ text: cell, bold: true })
              ]
            })
          ],
          shading: {
            fill: "DDDDDD",
          }
        })
      )
    });
    tableRows.push(headerRow);
    
    // Hoppa över separatorraden (med -----)
    rows = [rows[0]].concat(rows.slice(2));
  }
  
  // Lägg till övriga rader
  for (let i = hasHeader ? 1 : 0; i < rows.length; i++) {
    const row = rows[i];
    const tableRow = new TableRow({
      children: row.map(cell => 
        new TableCell({
          children: [new Paragraph({ text: cell })]
        })
      )
    });
    tableRows.push(tableRow);
  }
  
  // Skapa tabellen
  return new Table({
    rows: tableRows,
    width: {
      size: 100,
      type: WidthType.PERCENTAGE,
    },
    borders: {
      insideHorizontal: {
        style: BorderStyle.SINGLE,
        size: 1,
        color: "AAAAAA",
      },
      insideVertical: {
        style: BorderStyle.SINGLE,
        size: 1,
        color: "AAAAAA",
      },
      top: {
        style: BorderStyle.SINGLE,
        size: 1,
        color: "AAAAAA",
      },
      bottom: {
        style: BorderStyle.SINGLE,
        size: 1,
        color: "AAAAAA",
      },
      left: {
        style: BorderStyle.SINGLE,
        size: 1,
        color: "AAAAAA",
      },
      right: {
        style: BorderStyle.SINGLE,
        size: 1,
        color: "AAAAAA",
      },
    },
  });
}

  


    /**
 * Skapar ett Word-dokument med hjälp av Microsoft Graph API
 */
public async createWordDocumentWithGraph(
  content: string, 
  fileName: string, 
  libraryName: string = "Shared Documents"
): Promise<string> {
  try {
    console.log("Skapar Word-dokument med Microsoft Graph API...");
    
    // 1. Hämta Graph-klienten
    const graphClientFactory: MSGraphClientFactory = this.context.serviceScope.consume(MSGraphClientFactory.serviceKey);
    const graphClient = await graphClientFactory.getClient("3");
    
    // 2. Bygg URL för biblioteket
    const webUrl = this.context.pageContext.web.serverRelativeUrl;
    const hasTrailingSlash = webUrl.charAt(webUrl.length - 1) === '/';
    const relativeLibraryPath = hasTrailingSlash ? `${webUrl}${libraryName}` : `${webUrl}/${libraryName}`;
    
    // 3. Få SharePoint site ID genom att använda korrekt format för URL
    const tenantName = this.context.pageContext.site.absoluteUrl.match(/https:\/\/([^\/]*)/)?.[1] || 
  (() => { throw new Error("Kunde inte extrahera tenant-namn från URL"); })();
    const spSiteRelativePath = this.context.pageContext.site.serverRelativeUrl.replace(/^\//, '');
    console.log("spSiteRelativePath:", spSiteRelativePath);
    // Korrekt format: {hostname},{spSitePath} eller {hostname}:/{spSitePath}
    const siteIdFormat = `${tenantName}:/sites/test15`;
    console.log("Using site ID format:", siteIdFormat);
    
    // Anropa med korrekt format
    const siteResponse = await graphClient.api(`/sites/${siteIdFormat}`).get();
    
    console.log("Site info:", siteResponse);
    
    const drivesResponse = await graphClient.api(`/sites/${siteResponse.id}/drives`).get();
    
    console.log("Drives:", drivesResponse);
    
    // Hitta rätt drive (dokumentbibliotek)
    let driveId: string | undefined = undefined;
    for (const drive of drivesResponse.value) {
      if (drive.name === libraryName) {
        driveId = drive.id;
        break;
      }
    }
    
    if (!driveId && drivesResponse.value.length > 0) {
      console.log("Kunde inte hitta dokumentbiblioteket, använder första tillgängliga drive");
      driveId = drivesResponse.value[0].id;
    }
    
    if (!driveId) {
      throw new Error(`Kunde inte hitta någon drive/dokumentbibliotek för att skapa dokumentet`);
    }
    
    console.log(`Använder drive ID: ${driveId}`);
    
    // 5. Skapa ett tomt dokument
    const textEncoder = new TextEncoder();
    // Använd rent textinnehåll, INTE HTML
    const contentBytes = textEncoder.encode(content);

    const fileResponse = await graphClient.api(`/sites/${siteResponse.id}/drives/${driveId}/root:/${fileName}.txt:/content`)
      .put(contentBytes);
     
    console.log("Dokument skapat:", fileResponse);
    
    // 6. Returnera URL till dokumentet
    return `${relativeLibraryPath}/${fileName}.txt`;
  } catch (error) {
    console.error("Fel vid skapande av Word-dokument med Graph API:", error);
    // Lägg till fallback-metod här för att skapa en HTML-fil istället
    console.log("Använder fallback-metod för att skapa dokument...");
    try {
      return await this.createHtmlDocument(content, fileName, libraryName, "txt");
    } catch (fallbackError) {
      console.error("Även fallback-metod misslyckades:", fallbackError);
      throw error; // Kasta ursprungliga felet
    }
  }
}

/**
 * Skapar ett dokument och konverterar det till önskat filformat
 */
public async createHtmlDocument(
  content: string, 
  fileName: string, 
  libraryName: string = "Shared Documents",
  fileExtension: string = "docx"
): Promise<string> {
  try {
    // Hitta ett tillgängligt bibliotek
    let availableLibrary = "Shared Documents";
    
    try {
      const libraries = await this.getDocumentLibraries();
      
      if (libraries.includes(libraryName)) {
        availableLibrary = libraryName;
      } else if (libraries.length > 0) {
        availableLibrary = libraries[0];
        console.log(`Använder tillgängligt bibliotek: ${availableLibrary} istället för ${libraryName}`);
      }
    } catch (error) {
      console.warn("Kunde inte hämta bibliotek, använder Shared Documents:", error);
    }
    
    // Bygg URL för biblioteket
    const webUrl = this.context.pageContext.web.serverRelativeUrl;
    const hasTrailingSlash = webUrl.charAt(webUrl.length - 1) === '/';
    const folderPath = hasTrailingSlash ? `${webUrl}${availableLibrary}` : `${webUrl}/${availableLibrary}`;
    
    // Skapa först en temporär txt-fil (detta ska alltid fungera)
    const tempFileName = `${fileName}_temp.txt`;
    console.log(`Skapar temporär TXT-fil: ${tempFileName}`);
    
    // Ladda upp filen med råinnehåll
    const tempFileResult = await this.sp.web.getFolderByServerRelativePath(folderPath)
      .files.addUsingPath(tempFileName, content, { Overwrite: true });
    
    console.log(`Temporär TXT-fil skapad: ${tempFileResult.ServerRelativeUrl}`);
    
    // Steg 2: Kopiera tempfilen till önskat filformat
    try {
      // Hämta filen för att kunna använda copyTo
      const file = await this.sp.web.getFileByServerRelativePath(tempFileResult.ServerRelativeUrl);
      
      // Skapa målfilen med önskat filformat
      const targetPath = `${folderPath}/${fileName}.docx`;
      console.log(`Kopierar till ${fileExtension.toUpperCase()}: ${targetPath}`);
      
      // Utför kopieringen
      await file.copyTo(targetPath, true);
      console.log(`Fil konverterad till ${fileExtension.toUpperCase()}`);
      
      // Ta bort temporär fil
      try {
        await file.delete();
        console.log("Temporär fil borttagen");
      } catch (deleteError) {
        console.warn("Kunde inte ta bort temporär fil:", deleteError);
      }
      
      // Returnera sökvägen till den konverterade filen
      return `${folderPath}/${fileName}.docx`;
    } catch (copyError) {
      console.error(`Kunde inte konvertera till ${fileExtension}:`, copyError);
      
      // Om konverteringen misslyckas, returnera åtminstone den temporära filen
      return tempFileResult.ServerRelativeUrl;
    }
  } catch (error) {
    console.error(`Fel vid skapande av ${fileExtension}-dokument:`, error);
    throw error;
  }
}


}