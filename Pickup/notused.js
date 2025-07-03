function rota()
{
  // make a list of all schools per day
  //var linha = 6;  // linha student e days
  var l_out = 5; //linha student and schools 
  var colDiaStart = 3; // coluna iniciando monday
  const student_days = SpreadsheetApp.getActive().getSheetByName('Students x days'); // sheet Students x days
  const firstBlankStudent = getFirstEmptyRow(SpreadsheetApp.getActive().getSheetByName('Students x days'),6);
  //const schools = SpreadsheetApp.getActive().getSheetByName('Schools') // sheet Schools
  const kids_schools = SpreadsheetApp.getActive().getSheetByName('Kids and Schools') 
  const staffList = SpreadsheetApp.getActive().getSheetByName('Staff') // sheet staff
  var indexDia = {
    "Monday": colDiaStart,
    "Tuesday": colDiaStart+1,
    "Wednesday": colDiaStart+2,
    "Thursday": colDiaStart+3,
    "Friday": colDiaStart+4,
  };
  

  const dia = student_days.getRange(1,2).getValue().toString(); // dia escolhido para rota
  const data = student_days.getRange(1,1).getValue();
  var kdDia ={
      "Monday": "KD_MON",
      "Tuesday":"KD_TUE",
      "Wednesday":"KD_WED",
      "Thursday":"KD_THU",
      "Friday":"KD_FRI",
  }
  //console.log(kdDia[dia])
  const kd = SpreadsheetApp.getActive().getSheetByName(kdDia[dia]) 

  //console.log(dia);
  var rotaDia = SpreadsheetApp.getActive().getSheetByName(dia); // nome da planilha do dia da rota
  //var qte_child = student_days.getRange(linha,1).getValue(); //contador 
  
  // limpa planilhas
  //var ultimaLinhaRota = rotaDia.getLastRow();
  kids_schools.getRange("A5:D500").clear();
  kids_schools.getRange("A5:D500").setFontColor("black");
  kids_schools.getRange("A5:D500").clearDataValidations();
  //rotaDia.getRange("D:E").clearDataValidations();
  //console.log(ultimaLinhaRota);
  //rotaDia.getRange(ultimaLinhaRota, 2,ultimaLinhaRota+100,3).clear();
  //
  kids_schools.getRange("A2").setValue(dia);
  kids_schools.getRange("D3").setValue(data);
  

  let children = student_days.getRange("A6:H" + firstBlankStudent).getValues(); //.getValues();   //getDataRange().getValues();
  //let estaPresente = [];
  //let estaAusente = [];
  let statusDia = []; 
  let childrenPresents = [];
  let childrenAusentes = [];
  let childrenN = [];
  let staff = [];
  //console.log(children.length);
  for (var i = 0; i<children.length; i++){
    //var nome = children[i] 
    statusDia[i] = children[i][indexDia[dia]-1]; //student_days.getRange(linha,indexDia[dia]).getValues().toString();
    //estaAusente[i] = children[i][indexDia[dia]-1]; //student_days.getRange(linha,indexDia[dia]).getValues().toString();
    //estaN[i]= 
    //switch (actualSheetName)
     //{
    //case "MONDAY":

    //break;
    switch (statusDia[i]){
      case "P" || "p":
        //console.log("child is Present today: " + children[i][0]);
        childrenPresents.push([children[i][0],children[i][1],children[i][7]]);
      break;
      case "A" || "a":
        //console.log("child is Absent today: " + children[i][0]);
        childrenAusentes.push([children[i][0],children[i][1]]);
      break;
      case "N" || "n":
        //console.log("child is None today: " + children[i][0]);
        childrenN.push([children[i][0],children[i][1]]);
      break;
    }
    //if (statusDia[i] == "P" || estaPresente[i] == "p") 
    //{
      //console.log(children[i]);

      //aNome.push(children[i][0]); //children[i][0]\
      //childrenPresents.push([children[i][0],children[i][1],children[i][7]]);

    //} else if (estaAusente[i] == "A" || estaAusente == "a") {
      //childrenAusentes.push([children[i][0],children[i][1]]);
      
    //} else if (estaAusente[i] == "N" || estaAusente == "n") {
      //childrenN.push([children[i][0],children[i][1]]);
    //}
  } 
  childrenPresents.push(['<DRIVER SEAT>','','']);
  childrenPresents.push(['<HELPER SEAT>','','']);
  //console.log(childrenAusentes);
  kids_schools.getRange(5,1,childrenPresents.length,3).setValues(childrenPresents);


   // Preenche DRIVER SEAT AND HELPER SEAT
  //var ultimaLinhaKids = kids_schools.getLastRow();
  //l_out = ultimaLinhaKids+1;
  //kids_schools.getRange(l_out,1).setValue('<DRIVER SEAT>');
  //l_out++;
  //kids_schools.getRange(l_out,1).setValue('<HELPER SEAT>');
  //l_out++;
  //l_out++;
  //kids_schools.getRange(l_out,1).setValue('<< Absents: >>');
  l_out = kids_schools.getLastRow() + 2;
  if (childrenAusentes.length > 0 ) {
    kids_schools.getRange("D2").setValue(childrenAusentes.length);
    kids_schools.getRange(l_out,2).setValue("<< ABSENTS >>");
    l_out++;
    kids_schools.getRange(l_out,2,childrenAusentes.length,2).setValues(childrenAusentes);
    kids_schools.getRange(l_out,2,childrenAusentes.length,4).setFontColor("red");
  }

  // build a data validation with staff names
  //var lastRowStaff = staffList.getLastRow();
  //var firstBlankRowStaff = getFirstEmptyRow(SpreadsheetApp.getActive().getSheetByName('Staff'));
  //var firstBlankRowKids = getFirstEmptyRow(SpreadsheetApp.getActive().getSheetByName('Kids and Schools'));
  staff = staffList.getRange(2,1,staffList.getLastRow(),3).getValues();
  //console.log(firstBlankRowKids);
  
  let partRangeHelpers = [];
 // let partRangeDrivers = [];
  //let staffPresent = [];
  //let staffAbsent = [];
  var p = 0;
  for (var i = 0; i<staff.length; i++){
    if (staff[i][2] == "Present" ){
      partRangeHelpers.push(staff[i][0])
      //if (staff[i][1] == "Driver"){
      //  partRangeDrivers.push(staff[i][0])
      //}
    }
  }
  let partRangeChildrenPresents = [] //childrenPresents.slice(0); 
  
  for (var i = 0; i<childrenPresents.length;i++){
    partRangeChildrenPresents.push(childrenPresents[i][0])
  }
  //partRangeChildrenPresents.push("<DRIVER SEAT>");
  //partRangeChildrenPresents.push("<HELPER SEAT>");
  //console.log(partRangeHelpers);
  //console.log(partRangeDrivers);
  var partRuleHelpers = SpreadsheetApp.newDataValidation().requireValueInList(partRangeHelpers).build();
  //var partRuleDrivers = SpreadsheetApp.newDataValidation().requireValueInList(partRangeDrivers).build();
  var partRuleChildrenPresent = SpreadsheetApp.newDataValidation().requireValueInList(partRangeChildrenPresents).build();
  //uLinha = childrenPresents.length+6;
  //kids_schools.getRange("D5:D" + uLinha).setDataValidation(partRuleHelpers);
  
  // procurar crianca (Present) no dia
  let vehiclesRota = rotaDia.getRange("A1:G" + rotaDia.getLastRow()).getValues();
  //var nomeChild = "";
  let arrBlankSpaces = [];
  for (var c = 0; c<childrenPresents.length-2;c++) {
    var foundIt = false;
    //var lfounded = 0;
    
    //var nomeChild = childrenPresents[index][0];
    for (var i = 0; i<vehiclesRota.length; i++) {
      if (vehiclesRota[i][0] >= 1) {
        var lPresent = i + 1;
        rotaDia.getRange(lPresent,3).setDataValidation(partRuleChildrenPresent);
        rotaDia.getRange(lPresent,5).setDataValidation(partRuleHelpers);
        if (vehiclesRota[i][2] == ""){
          arrBlankSpaces.push(lPresent);
        }
      }
      if (vehiclesRota[i][2] == childrenPresents[c][0]){
        lfounded = lPresent;
        foundIt = true;
      }
    }
  
    if (foundIt == false){ // nao encontrou na lista 
      var nomeChild = childrenPresents[c][0];
      for (var s = 0; s<arrBlankSpaces.length; s++) { // loop with all blank spaces
        rotaDia.getRange(arrBlankSpaces[s],3).setValue(childrenPresents[c][0]);
        //rotaDia.getRange(l,5).setDataValidation(partRuleHelpers);
      }  
    }
  }
}


