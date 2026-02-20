// Fonction pour remplir les données manquantes dans l'Inventaire
function fillInventoryData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName("Inventaire");
  const confSheet = ss.getSheetByName("Config");
  
  if (!invSheet || !confSheet) {
    Logger.log("Erreur: Feuilles manquantes");
    return;
  }
  
  // Récupérer les fréquences
  const confData = confSheet.getDataRange().getValues();
  let frequencies = {};
  for (let i = 1; i < confData.length; i++) {
    frequencies[confData[i][0].toUpperCase()] = parseInt(confData[i][1]) || 30;
  }
  
  // Remplir l'Inventaire
  const invData = invSheet.getDataRange().getValues();
  const today = new Date();
  
  for (let i = 1; i < invData.length; i++) {
    const cat = invData[i][0] ? invData[i][0].toUpperCase() : "";
    const lastDateStr = invData[i][2];
    
    // Colonne 3 (Prochain_Controle) - index 3
    if (!invData[i][3] && lastDateStr) {
      let lastDate = new Date(lastDateStr);
      const freqDays = (frequencies[cat] || 30) * 30; // Convertir mois en jours (approximé)
      let nextDate = new Date(lastDate);
      nextDate.setDate(nextDate.getDate() + freqDays);
      invSheet.getRange(i + 1, 4).setValue(nextDate);
    }
    
    // Colonne 4 (Statut) - index 4
    if (!invData[i][4]) {
      let status = "green";
      if (invData[i][3]) {
        let nextDate = new Date(invData[i][3]);
        let daysLeft = Math.floor((nextDate - today) / (1000 * 60 * 60 * 24));
        
        if (daysLeft < 0) status = "red";
        else if (daysLeft < 30) status = "orange";
      }
      invSheet.getRange(i + 1, 5).setValue(status);
    }
  }
  
  Logger.log("Inventaire rempli!");
}
