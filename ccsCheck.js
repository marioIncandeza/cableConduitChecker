var projectSchedule = 'trial.xlsx';  //path to schedule goes here
var acceptable = ['n/a', 'TBD', 'TRAY', '4/0', 'SPARE'];  //list of entries that are ignored
var XLSX = require('xlsx');
var workbook = XLSX.readFile(projectSchedule);
var conduit = workbook.Sheets['Conduit Schedule'];  //the name of the sheet which corresponds to the Conduit Schedule
var cable = workbook.Sheets['Cable Schedule'];  //the name of the sheet which corresponds to the Calbe Schedule
var cellAddress;
var cellRef;
var conduitNo = [];
var cablesIncluded = [];
var cableNo = [];
var cableConduitPath = [];
var flagArray = [];

//Puts the contents of the 'Conduit No.' and 'Cables Included' columns into arrays. Blank rows are excluded. The two arrays are of equal length. Values are strings.
var maxRow = Number(conduit['!ref'].split(':').pop().split('').filter(function(el){return el.match(/\d/)>0;}).join(''));
for (var x=2; x<maxRow; x++){
	cellAddress = {c: 0, r: x};     //Column A
	cellRef = XLSX.utils.encode_cell(cellAddress);
	if (conduit[cellRef]){
		conduitNo.push(String(conduit[cellRef].v));
		//Grab what's in column B
		cellAddress.c++;
		cellRef = XLSX.utils.encode_cell(cellAddress);
		if (conduit[cellRef]){
			cablesIncluded.push(String(conduit[cellRef].v));
		}
		else {
			cablesIncluded.push('n/a');
		}
	}
}

//Puts the contents of the 'Cable No.' and 'Conduit No.' columns into arrays. Blank rows are excluded. The two arrays are of equal length. Values are strings.
maxRow = Number(cable['!ref'].split(':').pop().split('').filter(function(el){return el.match(/\d/)>0;}).join(''));
for (var x=2; x<maxRow; x++){
	cellAddress = {c: 0, r: x};     //Column A
	cellRef = XLSX.utils.encode_cell(cellAddress);
	if (cable[cellRef]){
		cableNo.push(String(cable[cellRef].v));
		//Grab what's in column B
		cellAddress.c++;
		cellRef = XLSX.utils.encode_cell(cellAddress);
		if (cable[cellRef]){
			cableConduitPath.push(String(cable[cellRef].v));
		}
		else {
			cableConduitPath.push('n/a');
		}
	}
}

//Split elements of cableConduitPath
var splitPath = [];
var holder = [];
cableConduitPath.forEach( function(el,inx){ 
	holder = el.split(',');
	holder.forEach(function(piece,indx){return holder[indx]=piece.trim();});
	return splitPath[inx] = holder;
});
//Split elements of cablesIncluded
var splitIncluded = [];
holder = [];
cablesIncluded.forEach( function(el,inx){ 
	holder = el.split(',');
	holder.forEach(function(piece,indx){return holder[indx]=piece.trim();});
	return splitIncluded[inx] = holder;
});
//Cable Check: starts with cable and verifies conduit path
var matchFlag = 0;
for (var c=0; c<cableNo.length; c++){
	for (var p=0; p<splitPath[c].length; p++){
		matchFlag = 1;
		if (acceptable.includes(splitPath[c][p])){
			matchFlag = 0;
		}
		var counter = 0;
		while(counter < conduitNo.length){
			if (splitPath[c][p]==conduitNo[counter]){
				for (var k = 0; k < splitIncluded[counter].length; k++){
					if(cableNo[c]==splitIncluded[counter][k]){
						matchFlag = 0;
					}
				}
			}
			counter++;
		}
		if (matchFlag==1){
			flagArray.push(cableNo[c]);
		}
	}
}

//Conduit Check: starts with conduit and checks included cables
var matchFlag = 0;
for (var c=0; c<conduitNo.length; c++){
	for (var p=0; p<splitIncluded[c].length; p++){
		matchFlag = 1;
		if (acceptable.includes(splitIncluded[c][p])){
			matchFlag = 0;
		}
		var counter = 0;
		while(counter < cableNo.length){
			if (splitIncluded[c][p]==cableNo[counter]){
				for (var k = 0; k < splitPath[counter].length; k++){
					if(conduitNo[c]==splitPath[counter][k]){
						matchFlag = 0;
					}
				}
			}
			counter++;
		}
		if (matchFlag==1){
			flagArray.push(conduitNo[c]);
		}
	}
}

console.log(flagArray);



