// Fonction de diagnostic pour vérifier la structure des données
function diagnosticSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let report = "";
  
  // Vérifier INVENTAIRE
  report += "=== INVENTAIRE ===\n";
  const invSheet = ss.getSheetByName("Inventaire");
  if (!invSheet) {
    report += "ERREUR: Feuille 'Inventaire' introuvable!\n\n";
  } else {
    const invData = invSheet.getDataRange().getValues();
    report += "Nombre de lignes: " + invData.length + "\n";
    report += "En-têtes: " + invData[0].join(" | ") + "\n";
    report += "Première ligne de données: " + (invData[1] ? invData[1].join(" | ") : "VIDE") + "\n\n";
  }
  
  // Vérifier CONFIG
  report += "=== CONFIG ===\n";
  const confSheet = ss.getSheetByName("Config");
  if (!confSheet) {
    report += "ERREUR: Feuille 'Config' introuvable!\n\n";
  } else {
    const confData = confSheet.getDataRange().getValues();
    report += "Nombre de lignes: " + confData.length + "\n";
    report += "En-têtes: " + confData[0].join(" | ") + "\n";
    for (let i = 1; i < Math.min(5, confData.length); i++) {
      report += "Ligne " + i + ": " + confData[i].join(" | ") + "\n";
    }
    report += "\n";
  }
  
  // Vérifier HISTORIQUE
  report += "=== HISTORIQUE ===\n";
  const histSheet = ss.getSheetByName("Historique");
  if (!histSheet) {
    report += "ERREUR: Feuille 'Historique' introuvable!\n\n";
  } else {
    report += "Nombre de lignes: " + histSheet.getLastRow() + "\n\n";
  }
  
  // Test getData()
  report += "=== TEST getData() ===\n";
  try {
    const result = getData();
    if (result.success) {
      report += "✓ getData() fonctionne!\n";
      report += "Nombre d'items inventaire: " + result.data.inventory.length + "\n";
      report += "Catégories: " + result.data.categoriesOrder.join(", ") + "\n";
      report += "Dashboard categories: " + Object.keys(result.data.dashboard).join(", ") + "\n";
    } else {
      report += "✗ Erreur getData(): " + result.error + "\n";
    }
  } catch (e) {
    report += "✗ Exception getData(): " + e.toString() + "\n";
  }
  
  Logger.log(report);
  return report;
}
