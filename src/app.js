var table = document.getElementById("ssTable");//get table by ID and store it as constant

//Below is the code to create table rows and columns... Initially adding 9 rows and 9 columns
for (let i=0; i<9; i++) {
    var row = document.getElementById("ssTable").insertRow(-1); //Js insertRow function... Insert row range is [-1,0]
    for(let j=0; j<9; j++){
        if(j<27){
        var letter=String.fromCharCode("A".charCodeAt(0)+j-1); //Getting Alphabetic numbers using fromCHarCode and CharCode At functions
    }else{
        var letter="A"+String.fromCharCode("A".charCodeAt(0)+j-27);
    }
	 //Conditional Ternary operator to add cells/letters/numbers
        row.insertCell(-1).innerHTML= i&&j ? "<input id='"+ letter+i +"'  class='tableCells'/>" : i||letter;
    }
};

//FUNCTION TO INSERT NEW ROWS
var insertButton=document.getElementById("insert");//Get insert button and add Event listener to it

var insertRows=function(){
    var rows = document.getElementById("ssTable").getElementsByTagName("tr").length; //Get no.of rows
    var columns=document.getElementById('ssTable').rows[0].cells.length;  //Get no.of Columns
    var newRow = document.getElementById("ssTable").insertRow(-1);  //Insert New Column
        for(let k=0; k<columns; k++){
            if(k<27){
                var letter=String.fromCharCode("A".charCodeAt(0)+k-1);
            }else{
                var letter="A"+String.fromCharCode("A".charCodeAt(0)+k-27);
            }
			//Conditional Ternary operator to add cells/letters/numbers
                newRow.insertCell(-1).innerHTML= k ? "<input id='"+ letter+rows +"' class='tableCells'/>" : rows;
                addingEvents();
            }
        
};
insertButton.addEventListener('click',insertRows);

//FUNCTION TO INSERT NEW ROWS AT SELECTED POSITION
    let insertatButton=document.getElementById("insertat"); //Get insert button and add Event listener to it
	
	
var insertRowsAt=function(){
    let rowLength = document.getElementById("ssTable").getElementsByTagName("tr").length;//Get no.of rows
    let columns=document.getElementById('ssTable').rows[0].cells.length; //Get no.of Columns
    let rowNo=document.getElementById("insertRowNo").value;  //Get the row no at which row needs to be inserted
    let newRow = document.getElementById("ssTable").insertRow(rowNo); //Insert New Column at the specified position
   
   //Insert the row and columns at given postion 
    for(let k=0; k<columns; k++){
        if(k<27){
            var letter=String.fromCharCode("A".charCodeAt(0)+k-1);
        }else{
            var letter="A"+String.fromCharCode("A".charCodeAt(0)+k-27);
        }
            newRow.insertCell(-1).innerHTML= k ? "<input id='"+ letter+rowNo +"' class='tableCells'/>" : rowNo;
        }
    //Update cells and row no after the inserted column
    for(let i=rowNo; i<rowLength+1; i++){
        
        for(let k=0; k<columns; k++){
            if(k<27){
                var letter=String.fromCharCode("A".charCodeAt(0)+k-1);
            }else{
                var letter="A"+String.fromCharCode("A".charCodeAt(0)+k-27);
            }
			//Conditional Ternary operator to add cells/letters/numbers
            table.rows[i].cells[k].innerHTML= k ? "<input id='"+ letter+i +"' value='"+"'/>" : i;
            }
        }
};
insertatButton.addEventListener('click',insertRowsAt);

//FUNTION TO INSERT COLUMN AT LAST POSTION
let insertColumnButton=document.getElementById("insertColumn"); //Get insert button and add Event listener to it

let insertColumn=function(){
    let rows = document.getElementById("ssTable").getElementsByTagName("tr").length; //Get no.of rows
    let columns=document.getElementById('ssTable').rows[0].cells.length; //Get no.of columns

    for(let x=0; x<rows; x++){
        if(columns<27){
            var letter=String.fromCharCode("A".charCodeAt(0)+columns-1);
        }else{
            var letter="A"+String.fromCharCode("A".charCodeAt(0)+columns-27);
        }
		//Conditional Ternary operator to add cells/letters/numbers
        table.rows[x].insertCell(-1).innerHTML= x ? "<input id='"+ letter+x +"' class='tableCells'/>" : letter;
        addingEvents();
    }
};
insertColumnButton.addEventListener('click',insertColumn);

//FUNTION TO DELETE LAST ROW
let deleteRowButton=document.getElementById("deleteRow"); //GET Delete button and add event listener to it

var deleteRow=function(){
    let rowLength = document.getElementById("ssTable").getElementsByTagName("tr").length; //Get no.of rows
    table.deleteRow(rowLength-1); //Delete row using js deleterow function
};
deleteRowButton.addEventListener('click',deleteRow);

//FUNTION TO DELETE LAST column
let deleteColumnButton=document.getElementById("deleteColumn");

var deleteColumn=function(){
    let lastRow = document.getElementById("ssTable").getElementsByTagName("tr"); //get row count
    let columnLength=document.getElementById('ssTable').rows[0].cells.length; //get column count
    for(let y=0; y<lastRow.length; y++){
       table.rows[y].deleteCell(columnLength-1); //Delete last cell of each row using deleteCell function
    }
};
deleteColumnButton.addEventListener('click',deleteColumn);
//Adding Events to table cells
var allCells=document.getElementsByClassName("tableCells"); //Get all the input fields using class name

(window.addingEvents=function(){for(let z=0; z<allCells.length; z++){
    //Add onfocus event to all the input fields
    allCells[z].onfocus=function(e){
        //e.target.value = "="+formula;
        getFormulas(e);
    }
    //Add onblur event to all the input fields
  allCells[z].onblur=function(e){  
      calculateAll(); //onblur it calls another function which evaluates the excel function on each cell if there is one 
  };
};
})();

//This function evaluates the excel functions on all cells that have functions or assigns the same value 
(window.calculateAll=function(){
	   for(var z=0; z<allCells.length; z++){
		   //Get the cells that have strings starting with '=' or has attribute 'formula' which we are assigning 
        if (allCells[z].value.charAt(0) === "="|| allCells[z].hasAttribute('formula')) {
			//Store the string without '=' by suing substring
            if(allCells[z].value.charAt(0) === "="){
                var formula=allCells[z].value.substring(1);
            }else{
                var formula=allCells[z].getAttribute('formula');
            }
            
            //Get the values at the mentioned cells using regex
            var expression = formula.replace(/([A-Z]+[0-9]+)/g||/([A-Z][A-z]+[0-9]+)/g, "parseFloat(document.getElementById('$1').value)") ;
            //Evaluate the value
            allCells[z].value= isNaN(parseFloat(eval(expression))) ? "" : parseFloat(eval(expression));
            //Set an attribute
            allCells[z].setAttribute('formula',formula);
        } else {  allCells[z].value= isNaN(parseFloat(allCells[z].value)) ? allCells[z].value : parseFloat(allCells[z].value); }
    }; 
 })();
 
 //This function checks if the cell has any expressions and displays them on focus
 (window.getFormulas=function(x){
    for(var z=0; z<allCells.length; z++){
        if(allCells[z].hasAttribute('formula') && x.target.id==allCells[z].id){
          allCells[z].value="="+allCells[z].getAttribute('formula');
        }
    }
})();

 function downloadCsv(csv, filename) {
    var csvFile;
    var downloadLink;

    // CSV FILE
    csvFile = new Blob([csv], {type: "text/csv"});

    // Download link
    downloadLink = document.createElement("a");

    // File name
    downloadLink.download = filename;

    // We have to create a link to the file
    downloadLink.href = window.URL.createObjectURL(csvFile);

    // Make sure that the link is not displayed
    downloadLink.style.display = "none";

    // Add the link to your DOM
    document.body.appendChild(downloadLink);

    // Lanzamos
    downloadLink.click();
}

function export_table(html, filename) {
	var csv = [];
	var rows = document.querySelectorAll("table tr");
	
    for (var i = 1; i < rows.length; i++) {
		var row = [], cols = rows[i].querySelectorAll(".tableCells");
		
        for (var j = 0; j < cols.length; j++) 
            row.push(cols[j].value);
        
		csv.push(row.join(","));		
	}

    // Download CSV
    downloadCsv(csv.join("\n"), filename);
}

document.getElementById("downloadButton").addEventListener("click", function () {
    var html = document.querySelector("table").outerHTML;
	export_table(html, "spreadsheet.csv");
});
 