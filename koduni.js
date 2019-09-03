function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('Form Yönetimi')
    .addItem('Şablondan Kopyala', 'copyFromTemplateWithSelection')
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('Yetkiler')
        .addItem('Bölge Yetkilendir', 'arrangeRegionPerms')
        .addItem('Heyet Yetkilendir', 'arrangeBoardPerms')
    )
    .addSeparator()
    .addToUi()
}

var spSheet=SpreadsheetApp.getActive()
var selection = spSheet.getActiveSheet().getSelection().getActiveRange()
var talebeDataSheet



var tplId = "1oHdn34dZaZ6yteEImo3b1bwXTx6B3Zv8wJqSV0-28bc"
var tplName = "UNI-Tekamüle Geçiş İmtihanı Mülakat Formu"



function merge() {
  //İcmal doyası
  var mergeFile=DriveApp.getFileById("1PoYUjm99Gh4YnZftR3NytK9zzg8ND6czZ0TyJhlPHEU")
  var mergeFileSpSheet = SpreadsheetApp.open(mergeFile)
  var mergeFileFormSheet =mergeFileSpSheet.getSheetByName('U-FORMB')



  var values = selection.getValues()
  for (var j = 0; j < values.length; j++) {
    var ssId=values[j][0]
    Logger.log(values[j][0])
    var sourceFile=DriveApp.getFileById(ssId)
    var sourceFileSpSheet = SpreadsheetApp.open(sourceFile)
    var sourceFileFormSheet =sourceFileSpSheet.getSheetByName('U-FORM')

    var filter=sourceFileFormSheet.getFilter()
    if(filter)
      filter.remove()

      
    /*
    //En son kaydı bulma
    var lastRow;
    var bolge=sourceFileFormSheet.getRange("N1:N164").getValues()
    var name=sourceFileFormSheet.getRange("S1:S164").getValues()
    for(i=164;i>4;i--){
      if(bolge[i-1]!=""||name[i-1]!=""){
         lastRow=i
         values[j][2]=lastRow
         break;
      }
    }
    selection.setValues(values)
    */
    

      
    //Kaynak sayfayı hedef dökümana kopyala
    var srcLastRow=values[j][2]
    var mStartRow=values[j][3]
    var mEndRow=values[j][4]

    
    var mergeFileSourceSheet=sourceFileFormSheet.copyTo(mergeFileSpSheet)

    var sRange=mergeFileSourceSheet.getRange("N5:DQ"+srcLastRow)
    var mRange=mergeFileFormSheet.getRange("N" +mStartRow + ":DQ" + mEndRow)
    sRange.copyTo(mRange)

    mergeFileSpSheet.deleteSheet(mergeFileSourceSheet)

    
  }
}


function copyFromTemplateWithSelection2() {
  talebeDataSheet=spSheet.getSheetByName('talebeler')
  
  var values = selection.getValues()
  for (var i = 0; i < values.length; i++) {
    var regionName = values[i][0]
    var file = DriveApp.getFileById("1MZnfOnJTMbnM7ZohUk1Xx20X90RclHDm9Pdpu7iTijM").makeCopy("Tekamül Talebe Listesi - " + regionName)

    
    values[i][1] = file.getId()
    selection.setValues(values)
    
  }
  
}

function copyFromTemplateWithSelection() {
  talebeDataSheet=spSheet.getSheetByName('talebeler')
  
  var values = selection.getValues()
  for (var i = 0; i < values.length; i++) {
    var regionName = values[i][0]
    var file = DriveApp.getFileById(tplId).makeCopy(tplName + ' - ' + regionName)

    
    values[i][1] = file.getId()
    // fill sheet with list
    var list = getListOfRegion(regionName)
    if(list.length){
      var fileSpSheet = SpreadsheetApp.open(file)
      var fileSheet = fileSpSheet.getSheetByName('U-FORM')
      var listRange = fileSheet.getRange(5, 13, list.length, list[0].length)
      listRange.setValues(list)
    }
    selection.setValues(values)
    
  }
}


function getListOfRegion (regionName) {
  var data = talebeDataSheet.getDataRange().getValues()
  
  data = data.filter(function (row) {
    return row[4] === regionName
  })
  
  data.forEach(function (row, i) {
    row[0] = i+1
  })
  return data
}








function permRow (data) {
  return {
    region: data[0],
    regionEmails: data[1].split(','),
    boardEmails: data[2].split(','),
    docId: data[3],
  }
}

function arrangeRegionPerms () {
  var values = selection.getValues()
  var regions = values.map(permRow)
  regions.forEach(function (region) {
    var file = DriveApp.getFileById(region.docId)
    var editors = file.getEditors()
    editors.forEach(function (ced) {
      if (region.regionEmails.indexOf(ced.getEmail()) === -1) {
        file.removeEditor(ced.getEmail())
      }
    })
    region.regionEmails.forEach(function (email) {
      email && file.addEditor(email)
    })
  })
}

function arrangeBoardPerms () {
  var values = selection.getValues()
  var regions = values.map(permRow)
  regions.forEach(function (region) {
    var file = DriveApp.getFileById(region.docId)
    var editors = file.getEditors()
    editors.forEach(function (ced) {
      if (region.boardEmails.indexOf(ced.getEmail()) === -1) {
        file.removeEditor(ced.getEmail())
      }
    })
    region.boardEmails.forEach(function (email) {
      try{
      email && file.addEditor(email)
      Logger.log(email)
      
      
      }catch(err){
        Logger.log(err)
        throw err
      }
      
    })
  })
}



function dene(){
  var values=['ADANA', 'ADIYAMAN', 'AFYONKARAHİSAR', 'AĞRI', 'AKSARAY', 'AMASYA', 'ANKARA', 'ANTALYA', 'ARDAHAN', 'ARTVİN', 'AYDIN', 'BALIKESİR', 'BARTIN', 'BATMAN', 'BAYBURT', 'BİLECİK', 'BİNGÖL', 'BİTLİS', 'BOLU', 'BURDUR', 'BURSA', 'ÇANAKKALE', 'ÇANKIRI', 'ÇORUM', 'DENİZLİ', 'DİYARBAKIR', 'DÜZCE', 'EDİRNE', 'ELAZIĞ', 'ERZİNCAN', 'ERZURUM', 'ESKİŞEHİR', 'GAZİANTEP', 'GİRESUN', 'GÜMÜŞHANE', 'HAKKARİ', 'HATAY', 'IĞDIR', 'ISPARTA', 'İSTANBUL', 'İZMİR', 'KAHRAMANMARAŞ', 'KARABÜK', 'KARAMAN', 'KARS', 'KASTAMONU', 'KAYSERİ', 'KIBRIS', 'KIRIKKALE', 'KIRKLARELİ', 'KIRŞEHİR', 'KİLİS', 'KOCAELİ', 'KONYA', 'KÜTAHYA', 'MALATYA', 'MANİSA', 'MARDİN', 'MERSİN', 'MUĞLA', 'MUŞ', 'NEVŞEHİR', 'NİĞDE', 'ORDU', 'OSMANİYE', 'RİZE', 'SAKARYA', 'SAMSUN', 'SİİRT', 'SİNOP', 'SİVAS', 'ŞANLIURFA', 'ŞIRNAK', 'TEKİRDAĞ', 'TOKAT', 'TRABZON', 'TUNCELİ', 'UŞAK', 'VAN', 'YALOVA', 'YOZGAT', 'ZONGULDAK', 'ABHAZYA', 'ADIGEY', 'AFGANİSTAN', 'ALMANYA', 'AMERİKA', 'ARNAVUTLUK', 'AVUSTRALYA', 'AVUSTURYA', 'AZERBAYCAN', 'BAE', 'BAHREYN', 'BANGLADEŞ', 'BELÇİKA', 'BENİN', 'BEYAZ RUSYA', 'BİSSAU', 'BOSNA HERSEK', 'BREZİLYA', 'BULGARİSTAN', 'BURKİNA FASO', 'BURNEİ', 'BURONDİ', 'CEZAYİR', 'CUBUTİ', 'ÇAD', 'ÇEÇENİSTAN', 'ÇEK CUMHURİYETİ', 'ÇİN', 'DAĞISTAN', 'DANİMARKA', 'DEMOKRATİK KONGO', 'ENDONEZYA', 'ETYOPYA', 'FAS', 'FİJİ ADALARI', 'FİLDİŞİ SAHİLLERİ', 'FİLİPİNLER', 'FİNLANDİYA', 'FRANSA', 'GADON', 'GAMBİYA', 'GANA', 'GAZZE', 'GİNE', 'GİNE', 'GÜNEY AFRİKA CUMHURİYETİ', 'GÜNEY KORE', 'GÜRCİSTAN', 'HIRVATİSTAN', 'HİNDİSTAN', 'HOLLANDA', 'HONG KONG', 'IRAK', 'İNGİLTERE', 'İNGUŞETYA', 'İSPANYA', 'İSVEÇ', 'İSVİÇRE', 'İTALYA', 'JAPONYA', 'KABARDAY BALKARYA', 'KALMUKYA', 'KAMBOÇYA', 'KAMERUN', 'KANADA', 'KARAÇAY ÇERKEZYA', 'KARADAĞ', 'KATAR', 'KAZAKİSTAN', 'KENYA', 'KIRGIZİSTAN', 'KONGO', 'KOSOVA', 'KUVEYT', 'LAOS', 'LESOTHO', 'LİBERYA', 'LİBYA', 'LÜBNAN', 'MACARİSTAN', 'MADAGASKAR', 'MAKODENYA', 'MALAVİ', 'MALDİVLER', 'MALEZYA', 'MALİ', 'MISIR', 'MOĞOLİSTAN', 'MOLDOVA', 'MORİTANYA', 'MOZAMBİK', 'MYANMAR', 'NEPAL', 'NİJER', 'NİJERYA', 'NORVEÇ', 'OSETYA', 'ÖZBEKİSTAN', 'PAKİSTAN', 'POLONYA', 'ROMANYA', 'RUANDA', 'RUSYA', 'SENEGAL', 'SIRBİSTAN', 'SİERRALEONE', 'SİNGAPUR', 'SLOVENYA', 'SOMALİ', 'SRİLANKA', 'SUDAN', 'SURİYE', 'SUUDİ ARABİSTAN', 'SÜRİNAM', 'TACİKİSTAN', 'TANZANYA', 'TAYLAND', 'TOGO', 'TUNUS', 'TÜRKMENİSTAN', 'UGANDA', 'UKRAYNA', 'UMMAN', 'ÜRDÜN', 'VENEZUELLA', 'VİETNAM', 'YEMEN', 'YENİ ZELLANDA', 'YUNANİSTAN', 'ZAMBİA', 'ZANZİBAR', 'ZİMBABVE']
  
  var rule=SpreadsheetApp.newDataValidation().requireValueInList(values, true).build()
  spSheet.getActiveSheet().getRange("E2:E4").setDataValidation(rule)
  
  var values = selection.getValues()
  for (var i = 0; i < values.length; i++) {
    var fileId = values[i][0]
    var file = DriveApp.getFileById(fileId)
    var fileSpSheet = SpreadsheetApp.open(file)
    var fileSheet = fileSpSheet.getSheetByName('U-FORM')
    fileSheet.getRange("R5:R164").setDataValidation(rule)
    fileSheet.getRange("AF5:AH164").setDataValidation(rule)

       
    
  }
  
  /*
  var targetFile=DriveApp.getFileById("1CmKdBZw1YGg6qCRl3qtXJCuy-Q-PmDxQH08YwgYELv4")
  var targetFileSpSheet = SpreadsheetApp.open(targetFile)
  var targetFileFormSheet =targetFileSpSheet.getSheetByName('FORM')  
  
  //templateFileFormSheet.copyTo(targetFileSpSheet)
  var targetFileSourceSheet =targetFileSpSheet.getSheetByName("source sayfasının kopyası")
  //targetFileSourceSheet.insertRows(1, 2)
  //targetFileSourceSheet.insertColumnsAfter(targetFileFormSheet.getLastColumn(),28)
  //targetFileSourceSheet.getRange("CR:DS").copyTo(targetFileFormSheet.getRange("CR:DS"))   
  targetFileFormSheet.setRowHeight(1, 36)
  targetFileFormSheet.setRowHeight(2, 48)
  
  var range=targetFileFormSheet.getRange("CR:DS")
  for (var i = range.getColumn(); i <= range.getLastColumn(); i++) {
    targetFileFormSheet.setColumnWidth(i, targetFileSourceSheet.getColumnWidth(i))
  }
  targetFileSpSheet.deleteSheet(targetFileSourceSheet)
  */
}