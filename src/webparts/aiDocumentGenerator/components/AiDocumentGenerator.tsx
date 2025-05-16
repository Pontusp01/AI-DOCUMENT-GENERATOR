import * as React from 'react';
import { useState, useEffect } from 'react';
import { TextField } from '@fluentui/react/lib/TextField';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { WordDocumentGenerator, WordGenerationRequest } from '../../../services/WordDocumentGenerator';

export interface IAiDocumentGeneratorProps {
  context: WebPartContext;
  displayMode: number;
  updateProperty: (value: string) => void;
}

export const AiDocumentGenerator: React.FC<IAiDocumentGeneratorProps> = (props) => {
  const [userPrompt, setUserPrompt] = useState<string>('');
  const [documentName, setDocumentName] = useState<string>('');
  const [templates, setTemplates] = useState<IDropdownOption[]>([]);
  const [selectedTemplate, setSelectedTemplate] = useState<string | undefined>(undefined);
  const [referenceMaterials, setReferenceMaterials] = useState<IDropdownOption[]>([]);
  const [selectedReferences, setSelectedReferences] = useState<string[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [isGenerating, setIsGenerating] = useState<boolean>(false);
  const [message, setMessage] = useState<{ text: string; type: MessageBarType } | null>(null);
  const [generatedDocUrl, setGeneratedDocUrl] = useState<string | null>(null);

  // Skapa wordGenerator när komponenten renderas
  const wordGenerator = new WordDocumentGenerator(props.context);

  // Ladda mallar och referensmaterial när komponenten monteras
  useEffect(() => {
    const loadData = async () => {
      try {
        setIsLoading(true);
        setMessage(null);
        
        // Försök ladda mallar
        try {
          const templateDocs = await wordGenerator.getAvailableTemplates();
          setTemplates(templateDocs.map((doc:any) => ({
            key: doc.url,
            text: doc.name
          })));
        } catch (error) {
          console.error("Fel vid laddning av mallar:", error);
          // Sätt templates till tom array om det misslyckas
          setTemplates([]);
        }
        
        // Försök ladda referensmaterial
        try {
          const refMaterials = await wordGenerator.getReferenceMaterials();
          setReferenceMaterials(refMaterials.map(doc => ({
            key: doc.url,
            text: `${doc.name} (${doc.contentType})`
          })));
        } catch (error) {
          console.error("Fel vid laddning av referensmaterial:", error);
          // Sätt referensmaterial till tom array om det misslyckas
          setReferenceMaterials([]);
        }
      } catch (error) {
        console.error("Fel vid datainläsning:", error);
        setMessage({
          text: `Fel vid inläsning av data: ${error instanceof Error ? error.message : "Okänt fel"}`,
          type: MessageBarType.error
        });
      } finally {
        setIsLoading(false);
      }
    };
    
    loadData();
  }, []);

  const handlePromptChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
    setUserPrompt(newValue || '');
  };

  const handleDocumentNameChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
    setDocumentName(newValue || '');
  };

  const handleTemplateChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    setSelectedTemplate(option ? option.key as string : undefined);
  };

  const handleReferenceToggle = (item: IDropdownOption, isSelected: boolean) => {
    if (isSelected) {
      setSelectedReferences([...selectedReferences, item.key as string]);
    } else {
      setSelectedReferences(selectedReferences.filter(key => key !== item.key));
    }
  };

  const generateDocument = async () => {
    // Validera inmatningar
    if (!userPrompt.trim()) {
      setMessage({
        text: "Vänligen ange en beskrivning av vad du vill generera.",
        type: MessageBarType.error
      });
      return;
    }
    
    if (!documentName.trim()) {
      setMessage({
        text: "Vänligen ange ett namn för det nya dokumentet.",
        type: MessageBarType.error
      });
      return;
    }
    
    // Förbered begäran
    const request: WordGenerationRequest = {
      userPrompt,
      documentName,
      selectedTemplateUrl: selectedTemplate,
      referenceUrls: selectedReferences
    };
    
    setIsGenerating(true);
    setMessage(null);
    
    try {
      const result = await wordGenerator.generateDocument(request);
      
      if (result.success && result.documentUrl) {
        setGeneratedDocUrl(result.documentUrl);
        setMessage({
          text: "Dokument har skapats framgångsrikt!",
          type: MessageBarType.success
        });
      } else {
        throw new Error(result.error || "Ett fel uppstod vid generering av dokumentet.");
      }
    } catch (error) {
      console.error("Fel vid generering av dokument:", error);
      setMessage({
        text: `Fel vid dokumentgenerering: ${error instanceof Error ? error.message : "Okänt fel"}`,
        type: MessageBarType.error
      });
    } finally {
      setIsGenerating(false);
    }
  };

  const stackTokens = { childrenGap: 15 };

  if (isLoading) {
    return (
      <Stack tokens={stackTokens}>
        <Spinner label="Laddar mallar och referensmaterial..." size={SpinnerSize.large} />
      </Stack>
    );
  }

  return (
    <Stack tokens={stackTokens} styles={{ root: { maxWidth: 800 } }}>
      <Text variant="xxLarge">AI Document Generator</Text>
      <Text>Skapa Word-dokument med hjälp av Azure OpenAI</Text>
      
      {message && (
        <MessageBar
          messageBarType={message.type}
          isMultiline={false}
          dismissButtonAriaLabel="Stäng"
        >
          {message.text}
        </MessageBar>
      )}
      
      <TextField
        label="Beskriv dokumentet du vill skapa"
        multiline
        rows={3}
        value={userPrompt}
        onChange={handlePromptChange}
        placeholder="Till exempel: 'Skapa en liknande offert för vår nya kund Företag AB'"
        required
      />
      
      <TextField
        label="Dokumentnamn"
        value={documentName}
        onChange={handleDocumentNameChange}
        placeholder="Ange ett namn för det nya dokumentet"
        required
      />
      
      <Dropdown
        label="Välj mall (valfritt)"
        options={templates}
        selectedKey={selectedTemplate}
        onChange={handleTemplateChange}
        placeholder="Välj en Word-mall att använda"
      />
      
      <Dropdown
        label="Välj referensmaterial (valfritt)"
        multiSelect
        options={referenceMaterials}
        selectedKeys={selectedReferences}
        onChange={(event, item) => item && handleReferenceToggle(item, !!item.selected)}
        placeholder="Välj sidor eller dokument som AI ska analysera"
      />
      
      <Stack horizontal tokens={{ childrenGap: 10 }}>
        <PrimaryButton 
          text="Generera dokument" 
          onClick={generateDocument} 
          disabled={isGenerating || !userPrompt.trim() || !documentName.trim()}
        />
        <DefaultButton 
          text="Återställ" 
          onClick={() => {
            setUserPrompt('');
            setDocumentName('');
            setSelectedTemplate(undefined);
            setSelectedReferences([]);
            setGeneratedDocUrl(null);
            setMessage(null);
          }} 
          disabled={isGenerating}
        />
      </Stack>
      
      {isGenerating && (
        <Spinner label="Genererar dokument..." size={SpinnerSize.large} />
      )}
      
      {generatedDocUrl && (
        <Stack>
          <Text>Ditt dokument har skapats:</Text>
          <DefaultButton
            text="Öppna dokument"
            href={`${props.context.pageContext.web.absoluteUrl}${generatedDocUrl}`}
            target="_blank"
          />
        </Stack>
      )}
    </Stack>
  );
};

export default AiDocumentGenerator;