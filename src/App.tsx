import { useEffect, useState, useRef, ChangeEvent } from 'react';
import QRCode from 'qrcode';
import { QrCode, Type, Image as ImageIcon, AlertCircle, CheckCircle2, Loader2, Settings2, Palette, ShieldCheck, Upload, X, Download, HelpCircle, ExternalLink } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

declare const Office: any;

type ErrorCorrectionLevel = 'L' | 'M' | 'Q' | 'H';

export default function App() {
  const [isOfficeInitialized, setIsOfficeInitialized] = useState(false);
  const [selectedText, setSelectedText] = useState<string>('');
  const [qrDataUrl, setQrDataUrl] = useState<string | null>(null);
  const [showOptions, setShowOptions] = useState(false);
  const [showInstallHelp, setShowInstallHelp] = useState(false);
  const [isAutoOpenEnabled, setIsAutoOpenEnabled] = useState(false);
  
  // Customization State
  const [errorCorrectionLevel, setErrorCorrectionLevel] = useState<ErrorCorrectionLevel>('M');
  const [fgColor, setFgColor] = useState('#000000');
  const [bgColor, setBgColor] = useState('#ffffff');
  const [logo, setLogo] = useState<string | null>(null);
  
  const [status, setStatus] = useState<{ type: 'idle' | 'loading' | 'success' | 'error'; message: string }>({
    type: 'idle',
    message: '',
  });

  const fileInputRef = useRef<HTMLInputElement>(null);

  const APP_URL = "https://ais-dev-6uerk3ttuve6n6aqe4iys3-70080317550.europe-west3.run.app";

  const downloadManifest = () => {
    const manifestXml = `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp 
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" 
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" 
  xsi:type="TaskPaneApp">
  <Id>d3f1a2b4-c5e6-4f7g-8h9i-0j1k2l3m4n5o</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Word QR Generator</ProviderName>
  <DefaultLocale>de-DE</DefaultLocale>
  <DisplayName DefaultValue="QR Code Generator" />
  <Description DefaultValue="Wandelt markierten Text in Word in einen QR-Code um und fügt ihn als Bild ein." />
  <IconUrl DefaultValue="${APP_URL}/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="${APP_URL}/icon-64.png" />
  <SupportUrl DefaultValue="${APP_URL}" />
  <AppDomains>
    <AppDomain>${APP_URL}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Document" />
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="${APP_URL}" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Document">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url" />
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="Commands.Group">
                <Label resid="Commands.GroupLabel" />
                <Control xsi:type="Button" id="Commands.TaskpaneButton">
                  <Label resid="Commands.TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="Commands.TaskpaneButton.Label" />
                    <Description resid="Commands.TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16" />
                    <bt:Image size="32" resid="Icon.32x32" />
                    <bt:Image size="80" resid="Icon.80x80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>MyTaskPaneID</TaskpaneId>
                    <SourceLocation resid="Commands.Url" />
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" resid="Icon.16x16" />
        <bt:Image id="Icon.32x32" resid="Icon.32x32" />
        <bt:Image id="Icon.80x80" resid="Icon.80x80" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Commands.Url" DefaultValue="${APP_URL}" />
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="${APP_URL}" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GetStarted.Title" DefaultValue="QR Generator bereit!" />
        <bt:String id="Commands.GroupLabel" DefaultValue="QR Tools" />
        <bt:String id="Commands.TaskpaneButton.Label" DefaultValue="QR Generator" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="Öffnen Sie den QR Generator, um markierten Text umzuwandeln." />
        <bt:String id="Commands.TaskpaneButton.Tooltip" DefaultValue="Öffnet das QR-Code Generator Panel." />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>`;

    const blob = new Blob([manifestXml], { type: 'text/xml' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'manifest.xml';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  useEffect(() => {
    // @ts-ignore
    if (typeof Office !== 'undefined') {
      // @ts-ignore
      Office.onReady((info) => {
        if (info.host === Office.HostType.Word) {
          setIsOfficeInitialized(true);
          // Check if auto-open is already set
          // @ts-ignore
          const autoOpen = Office.context.document.settings.get("Office.AutoShowTaskpaneWithDocument");
          setIsAutoOpenEnabled(!!autoOpen);
        }
      });
    }
  }, []);

  const toggleAutoOpen = () => {
    if (!isOfficeInitialized) return;
    const newValue = !isAutoOpenEnabled;
    setIsAutoOpenEnabled(newValue);
    
    // @ts-ignore
    Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", newValue);
    // @ts-ignore
    Office.context.document.settings.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        setStatus({ 
          type: 'success', 
          message: newValue 
            ? 'Dokument wird nun automatisch mit dem Add-in geöffnet.' 
            : 'Automatisches Öffnen deaktiviert.' 
        });
        setTimeout(() => setStatus({ type: 'idle', message: '' }), 3000);
      } else {
        setStatus({ type: 'error', message: 'Fehler beim Speichern der Einstellung.' });
      }
    });
  };

  const getSelection = async () => {
    if (!isOfficeInitialized) {
      setStatus({ type: 'error', message: 'Office ist nicht initialisiert oder wird nicht in Word ausgeführt.' });
      return;
    }

    setStatus({ type: 'loading', message: 'Text wird abgerufen...' });
    
    // @ts-ignore
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const text = result.value.trim();
        if (text) {
          setSelectedText(text);
          generateQR(text);
        } else {
          setStatus({ type: 'error', message: 'Bitte markieren Sie zuerst Text in Word.' });
        }
      } else {
        setStatus({ type: 'error', message: 'Fehler beim Abrufen des Textes: ' + result.error.message });
      }
    });
  };

  const generateQR = async (text: string) => {
    try {
      setStatus({ type: 'loading', message: 'QR-Code wird generiert...' });
      
      const canvas = document.createElement('canvas');
      await QRCode.toCanvas(canvas, text, {
        width: 800, // Higher resolution for better quality
        margin: 2,
        errorCorrectionLevel: errorCorrectionLevel,
        color: {
          dark: fgColor,
          light: bgColor,
        },
      });

      if (logo) {
        const ctx = canvas.getContext('2d');
        if (ctx) {
          const img = new Image();
          img.src = logo;
          await new Promise((resolve, reject) => {
            img.onload = resolve;
            img.onerror = () => reject(new Error('Logo konnte nicht geladen werden.'));
          });

          // Logo size: max 20% of QR code to ensure readability
          const logoSize = canvas.width * 0.18;
          const x = (canvas.width - logoSize) / 2;
          const y = (canvas.height - logoSize) / 2;
          
          // Draw a background for the logo to separate it from the QR patterns
          ctx.fillStyle = bgColor;
          ctx.beginPath();
          const padding = 6;
          ctx.roundRect(x - padding, y - padding, logoSize + padding * 2, logoSize + padding * 2, 8);
          ctx.fill();
          
          ctx.drawImage(img, x, y, logoSize, logoSize);
        }
      }

      const url = canvas.toDataURL('image/png');
      setQrDataUrl(url);
      setStatus({ type: 'idle', message: '' });
    } catch (err: any) {
      console.error(err);
      let msg = 'Fehler bei der QR-Code-Generierung.';
      if (err.message && (err.message.includes('too big') || err.message.includes('too long'))) {
        msg = 'Der Text ist zu lang für diesen QR-Code. Versuchen Sie eine niedrigere Fehlerkorrektur (L) oder kürzeren Text.';
      } else if (err.message) {
        msg = err.message;
      }
      setStatus({ type: 'error', message: msg });
      setQrDataUrl(null);
    }
  };

  const handleLogoUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      if (file.size > 500000) { // 500KB limit
        setStatus({ type: 'error', message: 'Das Logo ist zu groß (max. 500KB).' });
        return;
      }
      const reader = new FileReader();
      reader.onload = (event) => {
        setLogo(event.target?.result as string);
        if (selectedText) generateQR(selectedText);
      };
      reader.readAsDataURL(file);
    }
  };

  const clearLogo = () => {
    setLogo(null);
    if (fileInputRef.current) fileInputRef.current.value = '';
    if (selectedText) generateQR(selectedText);
  };

  const updateOption = (updater: () => void) => {
    updater();
    // Re-generate if we already have text
    if (selectedText) {
      setTimeout(() => generateQR(selectedText), 0);
    }
  };

  const insertQR = async () => {
    if (!qrDataUrl || !isOfficeInitialized) return;

    setStatus({ type: 'loading', message: 'Bild wird eingefügt...' });
    
    // Remove the data:image/png;base64, prefix
    const base64Image = qrDataUrl.split(',')[1];

    // @ts-ignore
    Office.context.document.setSelectedDataAsync(
      base64Image,
      { coercionType: Office.CoercionType.Image },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          setStatus({ type: 'success', message: 'QR-Code erfolgreich eingefügt!' });
          setTimeout(() => setStatus({ type: 'idle', message: '' }), 3000);
        } else {
          setStatus({ type: 'error', message: 'Fehler beim Einfügen: ' + result.error.message });
        }
      }
    );
  };

  return (
    <div className="min-h-screen bg-slate-50 p-4 font-sans text-slate-900">
      <header className="mb-6 flex items-center justify-between border-b border-slate-200 pb-4">
        <div className="flex items-center gap-2">
          <div className="rounded-lg bg-indigo-600 p-2 text-white">
            <QrCode size={24} />
          </div>
          <div>
            <h1 className="text-lg font-bold leading-tight">QR Generator</h1>
            <p className="text-xs text-slate-500">Word 2024 Ready</p>
          </div>
        </div>
        <button 
          onClick={() => setShowInstallHelp(!showInstallHelp)}
          className={`p-2 rounded-full transition-colors ${showInstallHelp ? 'bg-indigo-100 text-indigo-600' : 'text-slate-400 hover:bg-slate-100'}`}
          title="Installation & Hilfe"
        >
          <HelpCircle size={20} />
        </button>
      </header>

      <main className="space-y-6">
        <AnimatePresence>
          {showInstallHelp && (
            <motion.section
              initial={{ height: 0, opacity: 0 }}
              animate={{ height: 'auto', opacity: 1 }}
              exit={{ height: 0, opacity: 0 }}
              className="overflow-hidden rounded-2xl bg-indigo-50 border border-indigo-100 p-5 mb-6"
            >
              <h2 className="text-sm font-bold text-indigo-900 mb-3 flex items-center gap-2">
                <Download size={16} /> Installation in Word 2024
              </h2>
              <ol className="text-xs text-indigo-800 space-y-3 list-decimal list-inside">
                <li>Klicken Sie unten auf <strong>"Manifest herunterladen"</strong>.</li>
                <li>Öffnen Sie <strong>Word 2024</strong>.</li>
                <li>Suchen Sie auf der Registerkarte <strong>Start</strong> oder <strong>Einfügen</strong> nach der Schaltfläche <strong>"Add-ins"</strong>.</li>
                <li>Klicken Sie auf <strong>"Add-ins"</strong> und suchen Sie nach <strong>"Meine Add-ins"</strong> oder direkt nach <strong>"Add-in hochladen"</strong>.</li>
                <li>Wählen Sie die heruntergeladene <code>manifest.xml</code> aus.</li>
              </ol>
              <button
                onClick={downloadManifest}
                className="mt-4 w-full flex items-center justify-center gap-2 py-2.5 bg-indigo-600 text-white rounded-xl text-xs font-bold hover:bg-indigo-700 transition-colors shadow-sm"
              >
                <Download size={14} /> Manifest herunterladen
              </button>
              <div className="mt-4 pt-4 border-t border-indigo-200 flex items-center justify-between text-[10px] text-indigo-600 font-medium">
                <span>Version 1.0.0</span>
                <a href={APP_URL} target="_blank" rel="noreferrer" className="flex items-center gap-1 hover:underline">
                  App URL <ExternalLink size={10} />
                </a>
              </div>
            </motion.section>
          )}
        </AnimatePresence>

        <section className="rounded-2xl bg-white p-5 shadow-sm border border-slate-100">
          <div className="mb-4 flex items-center justify-between">
            <h2 className="flex items-center gap-2 text-sm font-semibold text-slate-700">
              <Type size={16} />
              Textquelle
            </h2>
            <div className="flex gap-2">
              <button 
                onClick={() => setShowOptions(!showOptions)}
                className={`p-1.5 rounded-lg transition-colors ${showOptions ? 'bg-indigo-100 text-indigo-600' : 'text-slate-400 hover:bg-slate-100'}`}
                title="Optionen"
              >
                <Settings2 size={18} />
              </button>
              {!isOfficeInitialized && (
                <span className="rounded-full bg-amber-100 px-2 py-0.5 text-[10px] font-medium text-amber-700">
                  Vorschau
                </span>
              )}
            </div>
          </div>

          <AnimatePresence>
            {showOptions && (
              <motion.div
                initial={{ height: 0, opacity: 0 }}
                animate={{ height: 'auto', opacity: 1 }}
                exit={{ height: 0, opacity: 0 }}
                className="overflow-hidden mb-4 space-y-4 border-t border-slate-100 pt-4"
              >
                {/* Colors */}
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-1.5">
                    <label className="text-[10px] font-bold uppercase text-slate-400 flex items-center gap-1">
                      <Palette size={10} /> Vordergrund
                    </label>
                    <div className="flex items-center gap-2">
                      <input 
                        type="color" 
                        value={fgColor} 
                        onChange={(e) => updateOption(() => setFgColor(e.target.value))}
                        className="h-8 w-full cursor-pointer rounded border border-slate-200 p-0.5"
                      />
                    </div>
                  </div>
                  <div className="space-y-1.5">
                    <label className="text-[10px] font-bold uppercase text-slate-400 flex items-center gap-1">
                      <Palette size={10} /> Hintergrund
                    </label>
                    <div className="flex items-center gap-2">
                      <input 
                        type="color" 
                        value={bgColor} 
                        onChange={(e) => updateOption(() => setBgColor(e.target.value))}
                        className="h-8 w-full cursor-pointer rounded border border-slate-200 p-0.5"
                      />
                    </div>
                  </div>
                </div>

                {/* Error Correction */}
                <div className="space-y-1.5">
                  <label className="text-[10px] font-bold uppercase text-slate-400 flex items-center gap-1">
                    <ShieldCheck size={10} /> Fehlerkorrektur
                  </label>
                  <div className="flex gap-1">
                    {(['L', 'M', 'Q', 'H'] as ErrorCorrectionLevel[]).map((level) => (
                      <button
                        key={level}
                        onClick={() => updateOption(() => setErrorCorrectionLevel(level))}
                        className={`flex-1 py-1.5 text-xs font-medium rounded-lg border transition-all ${
                          errorCorrectionLevel === level 
                            ? 'bg-indigo-50 border-indigo-200 text-indigo-600' 
                            : 'bg-white border-slate-200 text-slate-500 hover:border-slate-300'
                        }`}
                      >
                        {level}
                      </button>
                    ))}
                  </div>
                  <p className="text-[9px] text-slate-400 italic">
                    H (High) erlaubt Logos, reduziert aber die Textkapazität.
                  </p>
                </div>

                {/* Logo Upload */}
                <div className="space-y-1.5">
                  <label className="text-[10px] font-bold uppercase text-slate-400 flex items-center gap-1">
                    <ImageIcon size={10} /> Logo (Zentrum)
                  </label>
                  <div className="flex items-center gap-2">
                    {logo ? (
                      <div className="flex items-center justify-between w-full bg-slate-50 rounded-lg p-2 border border-slate-200">
                        <div className="flex items-center gap-2 overflow-hidden">
                          <img src={logo} alt="Logo preview" className="h-6 w-6 rounded object-cover" />
                          <span className="text-xs text-slate-500 truncate">Logo ausgewählt</span>
                        </div>
                        <button onClick={clearLogo} className="text-slate-400 hover:text-red-500">
                          <X size={14} />
                        </button>
                      </div>
                    ) : (
                      <button 
                        onClick={() => fileInputRef.current?.click()}
                        className="w-full flex items-center justify-center gap-2 py-2 text-xs font-medium text-slate-600 bg-slate-50 border border-dashed border-slate-300 rounded-lg hover:bg-slate-100 transition-colors"
                      >
                        <Upload size={14} /> Logo hochladen
                      </button>
                    )}
                    <input 
                      type="file" 
                      ref={fileInputRef} 
                      onChange={handleLogoUpload} 
                      accept="image/*" 
                      className="hidden" 
                    />
                  </div>
                </div>

                {/* Template / Auto-Open Option */}
                <div className="space-y-1.5 pt-2 border-t border-slate-100">
                  <label className="text-[10px] font-bold uppercase text-slate-400 flex items-center gap-1">
                    <HelpCircle size={10} /> Vorlagen-Optionen
                  </label>
                  <div className="flex items-center justify-between bg-slate-50 rounded-lg p-2 border border-slate-200">
                    <span className="text-xs text-slate-600">Auto-Öffnen mit Dokument</span>
                    <button
                      onClick={toggleAutoOpen}
                      className={`relative inline-flex h-5 w-9 items-center rounded-full transition-colors focus:outline-none ${
                        isAutoOpenEnabled ? 'bg-indigo-600' : 'bg-slate-300'
                      }`}
                    >
                      <span
                        className={`inline-block h-3 w-3 transform rounded-full bg-white transition-transform ${
                          isAutoOpenEnabled ? 'translate-x-5' : 'translate-x-1'
                        }`}
                      />
                    </button>
                  </div>
                  <p className="text-[9px] text-slate-400 italic">
                    Wenn aktiviert, öffnet sich das Add-in automatisch, wenn dieses Dokument (oder eine daraus erstellte Vorlage) geöffnet wird.
                  </p>
                </div>
              </motion.div>
            )}
          </AnimatePresence>

          <button
            onClick={getSelection}
            disabled={status.type === 'loading'}
            className="w-full rounded-xl bg-indigo-600 px-4 py-3 text-sm font-medium text-white transition-all hover:bg-indigo-700 active:scale-[0.98] disabled:opacity-50 flex items-center justify-center gap-2"
          >
            {status.type === 'loading' && status.message.includes('Text') ? (
              <Loader2 className="animate-spin" size={18} />
            ) : (
              <QrCode size={18} />
            )}
            Markierten Text umwandeln
          </button>

          {selectedText && (
            <div className="mt-4 rounded-lg bg-slate-50 p-3">
              <p className="text-[10px] uppercase tracking-wider text-slate-400 font-bold mb-1">Inhalt:</p>
              <p className="text-sm text-slate-600 line-clamp-3 break-all italic">"{selectedText}"</p>
            </div>
          )}
        </section>

        <AnimatePresence mode="wait">
          {qrDataUrl && (
            <motion.section
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="rounded-2xl bg-white p-5 shadow-sm border border-slate-100 text-center"
            >
              <h2 className="mb-4 flex items-center gap-2 text-sm font-semibold text-slate-700 justify-center">
                <ImageIcon size={16} />
                Vorschau
              </h2>
              
              <div className="mx-auto mb-6 aspect-square w-48 overflow-hidden rounded-xl border border-slate-100 bg-white p-2 shadow-inner">
                <img src={qrDataUrl} alt="Generated QR Code" className="h-full w-full object-contain" />
              </div>

              <button
                onClick={insertQR}
                disabled={status.type === 'loading'}
                className="w-full rounded-xl bg-emerald-600 px-4 py-3 text-sm font-medium text-white transition-all hover:bg-emerald-700 active:scale-[0.98] disabled:opacity-50 flex items-center justify-center gap-2"
              >
                {status.type === 'loading' && status.message.includes('Bild') ? (
                  <Loader2 className="animate-spin" size={18} />
                ) : (
                  <CheckCircle2 size={18} />
                )}
                In Word einfügen
              </button>
            </motion.section>
          )}
        </AnimatePresence>

        {/* Status Messages */}
        <AnimatePresence>
          {status.message && (
            <motion.div
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className={`flex items-center gap-3 rounded-xl p-4 text-sm ${
                status.type === 'error' ? 'bg-red-50 text-red-700 border border-red-100' :
                status.type === 'success' ? 'bg-emerald-50 text-emerald-700 border border-emerald-100' :
                'bg-indigo-50 text-indigo-700 border border-indigo-100'
              }`}
            >
              {status.type === 'error' ? <AlertCircle size={18} /> : 
               status.type === 'success' ? <CheckCircle2 size={18} /> : 
               <Loader2 className="animate-spin" size={18} />}
              <p>{status.message}</p>
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      <footer className="mt-8 text-center">
        <p className="text-[10px] text-slate-400">
          Entwickelt für Microsoft Word
        </p>
      </footer>
    </div>
  );
}
