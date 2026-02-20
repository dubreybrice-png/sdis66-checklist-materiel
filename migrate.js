// Fonction de migration pour corriger la structure des formulaires
function migrateFormStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Récupérer les données des feuilles Contenu
  const categories = ["VLI", "SAC ISP", "SAC IADE"];
  let formsData = {};
  
  categories.forEach(cat => {
    const sheet = ss.getSheetByName("Contenu " + cat);
    if (!sheet) {
      Logger.log("Feuille 'Contenu " + cat + "' non trouvée");
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    let sections = [];
    
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue; // Skip vide
      
      const section = data[i][0]; // Section
      const item = data[i][1]; // Item
      const type = data[i][2] ? data[i][2].toLowerCase() : "texte";
      const def = data[i][3] || "";
      const position = data[i][4] || "";
      
      // Chercher ou créer la section
      let sec = sections.find(s => s.section === section);
      if (!sec) {
        sec = { section: section, position: position, items: [] };
        sections.push(sec);
      }
      
      sec.items.push({ name: item, type: type, def: def });
    }
    
    formsData[cat] = sections;
  });
  
  // Sauvegarder dans Properties
  const prop = PropertiesService.getScriptProperties();
  prop.setProperty("FORMS_JSON", JSON.stringify(formsData));
  
  Logger.log("Migration réussie! Formulaires stockés.");
  Logger.log(JSON.stringify(formsData, null, 2));
}
