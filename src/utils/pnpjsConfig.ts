// src/utils/pnpjsConfig.ts
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { LogLevel, PnPLogging } from "@pnp/logging";
import { Caching } from "@pnp/queryable";

// Skapa en variabel för att hålla vår konfigurerade SP-instans
let _sp: SPFI | null = null;

// Funktion för att få eller initiera SP-instansen
export const getSP = (context?: WebPartContext): SPFI => {
  if (context != null) {
    // Konfigurera SP-instans med context och loggning
    _sp = spfi().using(SPFx(context)).using(PnPLogging(LogLevel.Warning)).using(Caching());
  }
  
  // Returnera den konfigurerade instansen
  return _sp!;
};