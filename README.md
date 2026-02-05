// ==========================================
// GOOGLE APPS SCRIPT pour Je joue Jeux partage
// ==========================================
// Ce script permet à l'application web d'interagir avec Google Sheets
// Il gère l'ajout, la modification et la suppression de jeux

// Fonction principale qui reçoit les requêtes POST
function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = JSON.parse(e.postData.contents);
    
    let result;
    
    switch(data.action) {
      case 'add':
        result = addGame(sheet, data);
        break;
      case 'update':
        result = updateGame(sheet, data);
        break;
      case 'delete':
        result = deleteGame(sheet, data);
        break;
      default:
        result = { success: false, error: 'Action inconnue' };
    }
    
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ 
        success: false, 
        error: error.toString() 
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Fonction pour gérer les requêtes GET (lecture seule)
function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  // Convertir en JSON (ignorer la première ligne d'en-têtes)
  const games = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) { // Si l'ID existe
      games.push({
        id: data[i][0],
        nom: data[i][1],
        joueurs: data[i][2],
        isbn: data[i][3],
        statut: data[i][4],
        image: data[i][5],
        date_ajout: data[i][6]
      });
    }
  }
  
  return ContentService
    .createTextOutput(JSON.stringify({ success: true, games: games }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Ajouter un jeu
function addGame(sheet, data) {
  const newRow = [
    data.id,
    data.nom,
    data.joueurs,
    data.isbn,
    data.statut,
    data.image || '',
    data.date_ajout || new Date().toISOString()
  ];
  
  sheet.appendRow(newRow);
  
  return { 
    success: true, 
    message: 'Jeu ajouté avec succès',
    id: data.id
  };
}

// Mettre à jour un jeu (pour changer le statut)
function updateGame(sheet, data) {
  const allData = sheet.getDataRange().getValues();
  
  // Trouver la ligne du jeu
  for (let i = 1; i < allData.length; i++) {
    if (allData[i][0] == data.id) {
      // Mettre à jour le statut (colonne E = index 4)
      sheet.getRange(i + 1, 5).setValue(data.statut);
      
      return { 
        success: true, 
        message: 'Statut mis à jour',
        id: data.id
      };
    }
  }
  
  return { 
    success: false, 
    error: 'Jeu non trouvé' 
  };
}

// Supprimer un jeu
function deleteGame(sheet, data) {
  const allData = sheet.getDataRange().getValues();
  
  // Trouver la ligne du jeu
  for (let i = 1; i < allData.length; i++) {
    if (allData[i][0] == data.id) {
      sheet.deleteRow(i + 1);
      
      return { 
        success: true, 
        message: 'Jeu supprimé',
        id: data.id
      };
    }
  }
  
  return { 
    success: false, 
    error: 'Jeu non trouvé' 
  };
}

// Fonction de test pour vérifier que tout fonctionne
function testAddGame() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  const testData = {
    id: Date.now().toString(),
    nom: 'Test Jeu',
    joueurs: '2-4 joueurs',
    isbn: '1234567890',
    statut: 'disponible',
    image: '',
    date_ajout: new Date().toISOString()
  };
  
  const result = addGame(sheet, testData);
  Logger.log(result);
}
