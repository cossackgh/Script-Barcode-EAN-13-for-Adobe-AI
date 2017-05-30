// Создание штрихкода EAN-13 v 2.1 
// (с) 2010-2016. Ягупов Дмитрий
// www.za-vod.ru
// info@za-vod.ru


// Create Barcode EAN-13 v 2.1
// (с) 2010-2016. Dmitry Yagupov 
// www.za-vod.ru
// info@za-vod.ru


#target illustrator
var fos =  Folder.fs;
//Get fullPath to Script file
if (fos == 'Windows'){
var pathJSX=$.fileName;
pathJSX = pathJSX.substring(0,pathJSX.lastIndexOf("\\"));
pathJSX = replace_string(pathJSX,'\\','/');
pathJSX = '/'+replace_string(pathJSX,':','');
}
else{
    
    var pathJSX=Folder.myDocuments;
//var pathJSX=Folder.desktop;
	}
	var mm=2.834645669; //convert point to mm
	var docPreset = new DocumentPreset;
	docPreset.units = RulerUnits.Millimeters;
	docPreset.width = 210*mm;
	docPreset.height = 297*mm;
    //docPreset.title = "Test Document";
	docPreset.colorMode = DocumentColorSpace.CMYK;



    try{
var docRef = activeDocument;


            }
    catch(ex){
 //alert(ex);		
		}


if (!docRef) {
	docRef = app.documents.addDocument(DocumentColorSpace.CMYK,docPreset);		
	//docRef = app.documents.add();  
	}


var EANGroup = docRef.groupItems.add(); // добавляем группу
//Read  INI File
var myINIFile = new File(pathJSX+'/barcode.ini');
var openIni = myINIFile.open('e'); // 'e' read-write open file.


           if (myINIFile.length == 0) {
 //Write  default INI file
        myINIFile. writeln('0'); // Start Coord X               
        myINIFile. writeln('0'); // Start Coord Y              
        myINIFile. writeln('Y'); // Chk Scale              
        myINIFile. writeln('N'); // Chk layer              
        myINIFile. writeln('100'); // Scale
        myINIFile. writeln('N');  // Chk DopParam
        myINIFile. writeln('1'); //Pages
        myINIFile. writeln('1'); //N Column
        myINIFile. writeln('1'); //N Row
        myINIFile. writeln('1'); // Dist Column
        myINIFile. writeln('1'); // Dist Row
	    myINIFile. writeln('/c'); // Path to Save AI Files		
        myINIFile.close();          
        
           myINIFile = new File(pathJSX+'/barcode.ini');
           var openIni = myINIFile.open('e'); // 'e' read-write open file.

}
if (!openIni){
var newPosX = '0';           
var newPosY = '0';           
//Read Check New Layer
var newLayer = 'Y';           
//Read Check Scale 
var newCheckScale = 'N';           
//Read Scale volume
var newScale = '100';           
//Read Clone Barcode
var newCheckClone = 'N';           
// Read Pages
var Pages = '1';       //Количество страниц при многостраничном документе    
//Read Column
var newColumn = '1';           
//Read Row
var newRow = '1';           
//Read ColumnDistance
var newColumnDist = '1';           
//Read RowDistance
var newRowDist = '1';  
//Read Full Path To Save AI Files
var destFolder = '/';     
    }
else{
//Read position 
var newPosX = myINIFile. readln();           
var newPosY = myINIFile. readln();           
//Read Check New Layer
var newLayer = myINIFile. readln();           
//Read Check Scale 
var newCheckScale = myINIFile. readln();           
//Read Scale volume
var newScale = myINIFile. readln();           
//Read  Check Extendet Parameters
var newCheckClone = myINIFile. readln();           
// Read Pages
var Pages = myINIFile. readln();       //Количество страниц при многостраничном документе    
//Read Column
var newColumn = myINIFile. readln();           
//Read Row
var newRow = myINIFile. readln();           
//Read ColumnDistance
var newColumnDist = myINIFile. readln();           
//Read RowDistance
var newRowDist = myINIFile. readln();  
//Read Full Path To Save AI Files
var destFolder = myINIFile. readln(); 

myINIFile.close();          
}  
var destFolderEncode = Folder.decode(destFolder);

var blk = 0.33;
var blkD = blk*7;
var blkE= blk*3;
var blkC= blk*5;
var blkH = 22.85; //Height standart bar
var blkHE = blkH+1.65; // Height extendet bar
//var blkHE = 24.5; // Height extendet bar
var zX=5; // Начальный отступ по X
var zY=5; // Начальный отступ по Y
var XGr=0; // Координаты группы одного штрихкода.X
var YGr=0; // Координаты группы одного штрихкода.Y
var ColMatrix=1; // Количество колонок при генерации нескольких штрих-кодов
var RowMatrix=1; // Количество строк при генерации нескольких штрих-кодов
var DistRow=1; // Расстояние между колонками при генерации нескольких штрих-кодов
var DistCol=1; // Расстояние между строками при генерации нескольких штрих-кодов
var namePages='EAN13-'; //Префикс названия страниц
var stepToNext=0;
var pathPrefix = null; // Name Disk from Path to Barcode AI files in ExtDialog
var pathLastFolder = null; // Text Last Folder from Path to Barcode AI files in ExtDialog
var textDlgPathAI = null; // Text Full Path to Barcode AI files in ExtDialog


var tablEAN= new Array(10);

tablEAN[0]="AAAAAA";
tablEAN[1]="AABABB";
tablEAN[2]="AABBAB";
tablEAN[3]="AABBBA";
tablEAN[4]="ABAABB";
tablEAN[5]="ABBAAB";
tablEAN[6]="ABBBAA";
tablEAN[7]="ABABAB";
tablEAN[8]="ABABBA";
tablEAN[9]="ABBABA";

var EAN = "";
var nowEnter="";




// Set Zero point ruler on Document
//Set Page Size to A4
docRef.width=210*mm;
docRef.height = 297*mm;
docRef.units = RulerUnits.Millimeters;



var hDoc = docRef.height;
var wDoc = docRef.width;

docRef.rulerOrigin = Array(0, hDoc); // Zero point ruler to left-top corner 



// Set color values for the CMYK object
var barColor = new CMYKColor();
barColor.black = 100;
barColor.cyan = 0;
barColor.magenta = 0;
barColor.yellow = 0;


//*******************************************
// Create Dialog Window

var win = new Window ("dialog {text:'EAN-13 ver 2.1  ||  (c) 2016 za-vod.ru', preferredSize:[329, 383], \
  tPanel0:Panel {type:'tabbedpanel', preferredSize:[295, 354], \
    tTab0:Panel {type:'tab', text:'Default param', \
      coord:Panel {text:'Barcode placed to:', orientation:'row', alignment:['center', 'top'], \
        name1:Group { \
          s1:StaticText {text:'X:'}, \
          e1:EditText {text:'0', characters:6}, \
          s2:StaticText {text:'mm   Y:'}, \
          e2:EditText {text:'0', characters:6}, \
          s3:StaticText {text:'mm'}}}, \
      digit12:Panel {text:'EAN-13', \
        name2:Group {orientation:'row', \
          s:StaticText {text:' Enter 12 digit code:'}, \
          e:EditText {text:'', helpTip:'Enter first 12 digit from your EAN code', characters:12}, \
          sh:StaticText {text:'<?>'}}}, \
      dopparamonoff:Panel {text:'', visible:false, \
        progressTxt:StaticText {text:'Progress:'}, \
        progressSave:Progressbar {value:0, minvalue:0, maxvalue:100}}, \
      buttons:Group {orientation:'row', alignment:['center', 'center'], \
        okBtn:Button {text:'OK'}, \
        cancelBtn:Button {text:'Cancel'}}, \
      gFrame:Group {orientation:'row', alignment:['center', 'center']}}, \
    tTab1:Panel {type:'tab', text:'Extendet param', \
      dopparam:Panel {text:'', orientation:'column', \
        name3:Group {orientation:'column', \
          chkLayer:Checkbox {text:'Barcode to new layer \"EAN-13\" ', alignment:['left', ''], value:true}, \
          chkScale:Checkbox {text:'Barcode Scale to (80-120%) ', alignment:['left', ''], value:false}, \
          sScale:StaticText {text:'Only 80-120%'}, \
          eScale:EditText {text:'100', helpTip:'If you want scale barcode, enter scale parameter.', characters:6, enabled:false}}}, \
      heightBarcode:Group {orientation:'row', \
        sHeight:StaticText {text:'Height Barcode (min 10 mm): '}, \
        eHeight:EditText {text:'22.85', helpTip:'Enter height your barcode. By default: 22.85 mm', characters:6}, \
        sHeightMM:StaticText {text:'mm'}}, \
      pages:Group {orientation:'row', alignChildren:['center', 'center'], \
        sPages:StaticText {text:'Pages'}, \
        ePages:EditText {text:'1', helpTip:'Enter number pages', characters:4}}, \
      pathAI:Group {orientation:'column', alignChildren:['center', ''], \
        sPath:StaticText {text:'Change Folder', helpTip:'Path to save your AI files', characters:30}, \
        PathBtn:Button {text:'Select folder to Save AI'}}, \
      name4:Group {orientation:'row', alignChildren:['right', ''], \
        sColumn:StaticText {text:'Columns'}, \
        eColumn:EditText {text:'1', characters:4}, \
        sRow:StaticText {text:'Rows'}, \
        eRow:EditText {text:'1', characters:4}}, \
      name5:Group {orientation:'row', alignChildren:['right', ''], \
        sDistanceX:StaticText {text:'Beth Columns'}, \
        eDistanceX:EditText {text:'1', characters:4}, \
        sDistanceY:StaticText {text:'Beth Rows'}, \
        eDistanceY:EditText {text:'1', characters:4}, \
        sDistanceYmm:StaticText {text:'mm'}}, \
      buttons:Group {orientation:'row', alignment:['center', ''], \
        extokBtn:Button {text:'Save Ext Param'}}}}}");


var sScale = win.tPanel0.tTab1.dopparam.name3.sScale;
var gfx = sScale.graphics;
gfx.foregroundColor = gfx.newPen(gfx.PenType.SOLID_COLOR, [1, 0, 0, 1], 1);


win.tPanel0.tTab0.coord.name1.e1.text = newPosX;
win.tPanel0.tTab0.coord.name1.e2.text = newPosY;

// Colorise 
var colorEAN = win.tPanel0.tTab0.digit12.graphics;
var myBrush = colorEAN.newBrush(colorEAN.BrushType.SOLID_COLOR, [0.5, 0.5, 0.5, 1]);
colorEAN.backgroundColor = myBrush;
var g = win.tPanel0.tTab1.pathAI.sPath;
var gGraph = g.graphics;
gGraph.font = ScriptUI.newFont("Verdana","BOLD",11);


if (newLayer=='Y'){
win.tPanel0.tTab1.dopparam.name3.chkLayer.value= true;
}
else win.tPanel0.tTab1.dopparam.name3.chkLayer.value= false;

if (newCheckScale=='Y'){
win.tPanel0.tTab1.dopparam.name3.chkScale.value= true;
win.tPanel0.tTab1.dopparam.name3.eScale.text= newScale;
win.tPanel0.tTab1.dopparam.name3.eScale.enabled= true;
}
else {
win.tPanel0.tTab1.dopparam.name3.chkScale.value= false;
win.tPanel0.tTab1.dopparam.name3.eScale.text= '100';
win.tPanel0.tTab1.dopparam.name3.eScale.enabled= false;
}

if (newCheckClone=='Y'){
	//win.tPanel0.tTab0.dopparamonoff.chkExtParam.value = true;
  //win.tPanel0.tTab0.dopparamonoff.extparamBtn.enabled = true; 
 //  win.tPanel0.tTab1.dopparam.name3.chkMatrix.value= true; 
   win.tPanel0.tTab1.pages.ePages.enabled = true;
   win.tPanel0.tTab1.name4.eColumn.enabled = true;
   win.tPanel0.tTab1.name4.eRow.enabled = true;
   win.tPanel0.tTab1.name5.eDistanceX.enabled = true;
   win.tPanel0.tTab1.name5.eDistanceY.enabled = true;  
   win.tPanel0.tTab1.pages.ePages.text = Pages;

   textDlgPathAI = ''+ destFolderEncode; 
//===========
	if (textDlgPathAI.length > 40) {

        if (fos == 'Windows'){
		pathPrefix = textDlgPathAI.substring(1,2)+':';
                                        }
        else{            	
            pathPrefix = textDlgPathAI.substring(textDlgPathAI.indexOf('/',13),textDlgPathAI.indexOf('/',16));
            pathPrefix = '~' + pathPrefix;

            }
		pathPrefix= pathPrefix.toUpperCase();
		pathLastFolder = textDlgPathAI.substring(textDlgPathAI.lastIndexOf("/"));
		textDlgPathAI= 'Path to Save AI -> ' + pathPrefix + '...' + pathLastFolder;
		win.tPanel0.tTab1.pathAI.sPath.text = textDlgPathAI ;
        win.tPanel0.tTab1.pathAI.sPath.helpTip = destFolderEncode;
		}
	else{
                if (fos == 'Windows'){
		pathPrefix = textDlgPathAI.substring(1,2)+':';
		pathLastFolder = textDlgPathAI.substring(2);
                                        }
        else{
            pathPrefix = '~';
		pathLastFolder = textDlgPathAI.substring(15);
                }
		textDlgPathAI= 'Path to Save AI -> ' + pathPrefix +  pathLastFolder;		
		win.tPanel0.tTab1.pathAI.sPath.text = textDlgPathAI ;
        win.tPanel0.tTab1.pathAI.sPath.helpTip = destFolderEncode;
			}

//==========


   win.tPanel0.tTab1.name4.eColumn.text = newColumn;
   win.tPanel0.tTab1.name4.eRow.text = newRow;
   win.tPanel0.tTab1.name5.eDistanceX.text = newColumnDist;
   win.tPanel0.tTab1.name5.eDistanceY.text = newRowDist;     
    }
else{

   win.tPanel0.tTab1.pages.ePages.enabled = false;
   win.tPanel0.tTab1.name4.eColumn.enabled = false;
   win.tPanel0.tTab1.name4.eRow.enabled = false;
   win.tPanel0.tTab1.name5.eDistanceX.enabled = false;
   win.tPanel0.tTab1.name5.eDistanceY.enabled = false;   
   win.tPanel0.tTab1.pages.ePages.text = '1';   
   win.tPanel0.tTab1.name4.eColumn.text = '1';
   win.tPanel0.tTab1.name4.eRow.text = '1';
   win.tPanel0.tTab1.name5.eDistanceX.text = '1';
   win.tPanel0.tTab1.name5.eDistanceY.text = '1';      
    }



// Draw logo



var colorFr = win.tPanel0.tTab0.gFrame.graphics;
var btnFrame = win.tPanel0.tTab0.gFrame;
btnFrame.size = [130,150];
btnFrame.margins = [10,50,10,50];

var myFrBrush = colorFr.newBrush(colorFr.BrushType.SOLID_COLOR, [0.1, 0.4, 0, 1]);

btnFrame.LineColorContur = btnFrame.graphics.newPen(btnFrame.graphics.PenType.SOLID_COLOR, [1, 0, 0, 1],1 );
btnFrame.BrushColorRect = btnFrame.graphics.newBrush(btnFrame.graphics.BrushType.SOLID_COLOR, [1, 0, 0, 1]);

btnFrame.LineConturBlack = btnFrame.graphics.currentPath;
btnFrame.onDraw = customDrawBar;


btnFrame.text = "Barcode    EAN-13";  
btnFrame.textPen = btnFrame.graphics.newPen (btnFrame.graphics.PenType.SOLID_COLOR,[0,0,0,1], 1);
btnFrame.text2 = "http://za-vod.ru";  
btnFrame.textPen = btnFrame.graphics.newPen (btnFrame.graphics.PenType.SOLID_COLOR,[0,0,0,1], 1);

colorFr.backgroundColor = myFrBrush;

function customDrawBar(){   
    with( this ) {  
graphics.drawOSControl();  
graphics.rectPath(10,10,2,55);  
graphics.rectPath(14,10,2,55); 
graphics.rectPath(18,10,4,50); 
graphics.rectPath(24,10,2,50); 
graphics.rectPath(28,10,2,50); 
graphics.rectPath(36,10,2,50); 
graphics.rectPath(40,10,2,50); 
graphics.rectPath(44,10,2,50); 
graphics.rectPath(50,10,4,50); 
graphics.rectPath(56,10,2,50); 
graphics.rectPath(60,10,2,50); 
graphics.rectPath(66,10,2,55); 
graphics.rectPath(70,10,2,55); 
graphics.rectPath(74,10,2,50); 
graphics.rectPath(78,10,2,50); 
graphics.rectPath(84,10,2,50); 
graphics.rectPath(88,10,4,50); 
graphics.rectPath(94,10,2,50); 
graphics.rectPath(98,10,2,50); 
graphics.rectPath(102,10,2,50); 
graphics.rectPath(108,10,6,50); 
graphics.rectPath(118,10,2,50); 
graphics.rectPath(122,10,2,55); 
graphics.rectPath(126,10,2,55); 

graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, [0, 0, 0, 1])); 
graphics.drawString(text,textPen,18,62,graphics.font);
graphics.drawString(text2,textPen,18,75,graphics.font);


}} 

// End Draw logo



		      
		      newCheckClone = 'Y';			 
           win.tPanel0.tTab1.pages.ePages.enabled = true;            
           win.tPanel0.tTab1.name4.eColumn.enabled = true;
           win.tPanel0.tTab1.name4.eRow.enabled = true;
           win.tPanel0.tTab1.name5.eDistanceX.enabled = true;
           win.tPanel0.tTab1.name5.eDistanceY.enabled = true;
            




//OnClick Save btn
win.tPanel0.tTab1.buttons.extokBtn.onClick = function ExtParamSave(){
    var txtSaveExt = "Ext param: Column - ";
    txtSaveExt = txtSaveExt+win.tPanel0.tTab1.name4.eColumn.text +" \n Row - ";
    txtSaveExt = txtSaveExt+win.tPanel0.tTab1.name4.eRow.text;
    writeINI();

    }

// Get Path for saved AI files
win.tPanel0.tTab1.pathAI.PathBtn.onClick = function(){
var olddestFolder=destFolder;
	destFolder = Folder.selectDialog( 'Select folder for Save Barcode files.', destFolder);
    if (!destFolder){
        destFolder=olddestFolder;// Bad code
        }
    destFolderEncode = Folder.decode(destFolder);
	textDlgPathAI = ''+ destFolderEncode; 

	if (textDlgPathAI.length > 40) {

        if (fos == 'Windows'){
		pathPrefix = textDlgPathAI.substring(1,2)+':';
                                        }
        else{            	
            pathPrefix = textDlgPathAI.substring(textDlgPathAI.indexOf('/',13),textDlgPathAI.indexOf('/',16));
            pathPrefix = '~' + pathPrefix;

            }
		pathPrefix= pathPrefix.toUpperCase();
		pathLastFolder = textDlgPathAI.substring(textDlgPathAI.lastIndexOf("/"));
		textDlgPathAI= 'Save AI to Folder -> ' + pathPrefix + '...' + pathLastFolder;
		win.tPanel0.tTab1.pathAI.sPath.text = textDlgPathAI ;
		win.tPanel0.tTab1.pathAI.sPath.helpTip = destFolderEncode;
		}
	else{
                if (fos == 'Windows'){
		pathPrefix = textDlgPathAI.substring(1,2)+':';
		pathLastFolder = textDlgPathAI.substring(2);
                                        }
        else{
         pathPrefix = '~';
		pathLastFolder = textDlgPathAI.substring(15);
                }
		textDlgPathAI= 'Save AI to Folder-> ' + pathPrefix +  pathLastFolder;		
		win.tPanel0.tTab1.pathAI.sPath.text = textDlgPathAI ;   
		win.tPanel0.tTab1.pathAI.sPath.helpTip = destFolderEncode;
			}
	}

// Check If enter only digit 0-9
win.tPanel0.tTab0.digit12.name2.e.onChanging = function (){    
ChangeEANInput();
    }

// If Pages >1 Get Path for saved AI files
win.tPanel0.tTab1.pages.ePages.onChanging = function(){
	var chngPages = parseInt(win.tPanel0.tTab1.pages.ePages.text);
if ( chngPages <1){
		//destFolder = null;
		win.tPanel0.tTab1.pages.ePages.text = '1';	
	}
	}


function ChangeEANInput(){
 	nowEnter = win.tPanel0.tTab0.digit12.name2.e.text;
	var vPattern = /[^0-9]/;
	var noneD = /\D/g;
	var result = vPattern.test(nowEnter);

if (result == true)
{
	nowEnter = nowEnter.replace(noneD, "") ;
	win.tPanel0.tTab0.digit12.name2.e.text = nowEnter;
    alert('Only numbers are permitted for this field.');
}

	
	if ( nowEnter.length > 12) {
		alert('You enter more 12 digit');
		nowEnter = nowEnter.substring(0,12);
		win.tPanel0.tTab0.digit12.name2.e.text =  nowEnter;
		
		}
	
    var chk13 = SUM13(nowEnter);    

    EAN = nowEnter+chk13;	
	win.tPanel0.tTab0.digit12.name2.sh.text = chk13;   
    
    }


// Height  field onChange
win.tPanel0.tTab1.heightBarcode.eHeight.onChange = function ChangeHeghtInput(){
 	blkH = win.tPanel0.tTab1.heightBarcode.eHeight.text;
	var vPattern = /[^0-9.]/;
	var noneD = /\D/g;
	var result = vPattern.test(blkH);

if (result == true)
{
	blkH = blkH.replace(noneD, "") ;
	win.tPanel0.tTab1.heightBarcode.eHeight.text = blkH;
    alert('Only numbers are permitted for this field.');
}

	
	if ( blkH.length > 4) {
		alert('You enter more 4 digit');
		blkH = blkH.substring(0,4);
		win.tPanel0.tTab1.heightBarcode.eHeight.text =  blkH;
		}
	if (parseInt(blkH) < 10){
				alert('You enter less then 10 mm');
				win.tPanel0.tTab1.heightBarcode.eHeight.text =  '10';
				
		}
	
    
    }

// OK botton Click
win.tPanel0.tTab0.buttons.okBtn.onClick = function actionPlace() { 
    var enterDigits = win.tPanel0.tTab0.digit12.name2.e.text.length;
    var newLayer = win.tPanel0.tTab1.dopparam.name3.chkLayer.value;
    var enterScale = parseInt(win.tPanel0.tTab1.dopparam.name3.eScale.text);
    var ColMatrix=parseInt(win.tPanel0.tTab1.name4.eColumn.text);
    var RowMatrix=parseInt(win.tPanel0.tTab1.name4.eRow.text);
    var DistRow=parseInt(win.tPanel0.tTab1.name5.eDistanceX.text);
    var DistCol=parseInt(win.tPanel0.tTab1.name5.eDistanceX.text);
          Pages = parseInt(win.tPanel0.tTab1.pages.ePages.text);
         // alert("START Pages = "+ enterDigits);  
    var First12="";
    var GrHeight=0;
          //stepToNext++;
    var posXGroup = win.tPanel0.tTab0.coord.name1.e1.text;
          XGr = parseInt(posXGroup);    
    var posYGroup = win.tPanel0.tTab0.coord.name1.e2.text;    
          YGr = parseInt(posYGroup); 
	var FullPathToSave = null;
	var fileSaveBCode  = null;
	blkH= parseFloat(win.tPanel0.tTab1.heightBarcode.eHeight.text);
	blkHE = blkH +1.65; // Height extendet bar

    if ( win.tPanel0.tTab1.dopparam.name3.chkLayer.value == true) {
                chkLayer();
                                                                                }  
                                                                                        
            
    if ( enterDigits == 12) {   


        if (( enterScale < 80) || (enterScale >120))      // проверяем диапазон масштабирования 80-120%
            alert('Wrong Scale. Enter 80-120% only');        
        else {
           
             if (Pages > 1) {


                 //RowMatrix = RowMatrix/Pages;
                 // Show ProgressBar
                 win.tPanel0.tTab0.dopparamonoff.progressTxt.visible = true;
                 win.tPanel0.tTab0.dopparamonoff.progressSave.visible = true;                 
                 win.update();
                 
				 for (var p =1; p<=Pages; p++){

                     win.tPanel0.tTab0.dopparamonoff.progressSave.value = p/Pages*100; // update progressbar
                     win.update();
			  XGr = parseInt(posXGroup);
			  YGr = parseInt(posYGroup);
                docRef = app.documents.addDocument(DocumentColorSpace.CMYK,docPreset);	 

                EANGroup = docRef.groupItems.add(); // добавляем группу
                hDoc = docRef.height;
                wDoc = docRef.width;
                docRef.rulerOrigin = Array(0, hDoc);   
        if ( win.tPanel0.tTab1.dopparam.name3.chkLayer.value == true) {
                chkLayer();
                                                                                }             
                
             for ( var m=0; m<RowMatrix; m++){
                 for ( var n=0;n<ColMatrix;n++){
                  First12 =  EAN.substring(0,12);

        win.tPanel0.tTab0.digit12.name2.e.text  =  parseInt(First12)+stepToNext;
		stepToNext = 1;// Bad solution :(
        ChangeEANInput();

        CreatEAN(); // Рисуем штрихкод
        EANGroup.resize(enterScale,enterScale); // Масштабируем
        GrWidth=EANGroup.width/mm;  // Вычисляем ширину группы с одним штрихкодом      
        GrHeight=EANGroup.height/mm;  // Вычисляем высоту группы с одним штрихкодом              
        XGr=XGr+GrWidth+ parseInt(win.tPanel0.tTab1.name5.eDistanceX.text); // Координата X следующего блока штрихкода
        EANGroup = docRef.groupItems.add(); // добавляем группу
        }
        XGr = parseInt(posXGroup); // Координата X следующего блока штрихкода
        YGr=YGr+GrHeight+ parseInt(win.tPanel0.tTab1.name5.eDistanceY.text); // Координата Y следующего блока штрихкода
        }

	FullPathToSave = destFolder+'/'+namePages+EAN+'.ai';	
		// Create the file object to save to
	fileSaveBCode = new File( FullPathToSave);

    docRef.saveAs(fileSaveBCode);
	docRef.close();
                 } // End For Pages
                 } // End If Pages >1
             else {
             for ( var m=0; m<RowMatrix; m++){
                 for ( var n=0;n<ColMatrix;n++){
                  First12 =  EAN.substring(0,12);

        win.tPanel0.tTab0.digit12.name2.e.text  =  parseInt(First12)+stepToNext;
		stepToNext = 1; // Bad solution :(
        ChangeEANInput();
        CreatEAN(); // Рисуем штрихкод
        EANGroup.resize(enterScale,enterScale); // Масштабируем
        GrWidth=EANGroup.width/mm;  // Вычисляем ширину группы с одним штрихкодом      
        GrHeight=EANGroup.height/mm;  // Вычисляем высоту группы с одним штрихкодом              
        XGr=XGr+GrWidth+ parseInt(win.tPanel0.tTab1.name5.eDistanceX.text); // Координата X следующего блока штрихкода
        EANGroup = docRef.groupItems.add(); // добавляем группу
        } // End For Column
        XGr = parseInt(posXGroup); // Координата X следующего блока штрихкода
        YGr=YGr+GrHeight+ parseInt(win.tPanel0.tTab1.name5.eDistanceY.text); // Координата Y следующего блока штрихкода
        } //End For Row
                } // End else if Pages =1
            writeINI(); // Записываем INI файл        
            actionCanceled(); // Заканчиваем скрипт


               
                }
        }
    else 
    alert ('You do NOT Enter 12 digits');
    
}

//проверяем масштабирование
win.tPanel0.tTab1.dopparam.name3.chkScale.onClick = function addScale() {     
    if (win.tPanel0.tTab1.dopparam.name3.chkScale.value == true)    {
    win.tPanel0.tTab1.dopparam.name3.eScale.enabled = true;
    enterScale = parseInt(win.tPanel0.tTab1.dopparam.name3.eScale.text);
            }
    else {
    win.tPanel0.tTab1.dopparam.name3.eScale.enabled = false;    
    win.tPanel0.tTab1.dopparam.name3.eScale.text = '100';
    enterScale = 100;
    
    }
    }

win.tPanel0.tTab0.buttons.cancelBtn.onClick = function exitDlg() { 

	win.close();
    }


// Проверяем ввод только цифр  и диапазон 80-120%
win.tPanel0.tTab1.dopparam.name3.eScale.onChanging = function (){  
	var nowEnterScale = win.tPanel0.tTab1.dopparam.name3.eScale.text;
	var vPattern = /[^0-9]/;
	var noneD = /\D/g;
	var result = vPattern.test(nowEnterScale);

if (result == true)
{
	nowEnterScale = nowEnterScale.replace(noneD, "") ;
	win.tPanel0.tTab1.dopparam.name3.eScale.text = nowEnterScale;
    alert('Only numbers are permitted for this field.');
}

	
	if ( nowEnterScale.length > 3) {
		alert('You enter more 3 digit');
		nowEnterScale = nowEnterScale.substring(0,3);
		win.tPanel0.tTab1.dopparam.name3.eScale.text =  nowEnterScale;		
		}

    }

win.center(); 
win.show();


 function actionCanceled() { 

	win.close();
}

// Если нужен штрих-код на новом слое
function chkLayer(){
    //create layer "EAN-13" if exist
    try{
    var stL = docRef.layers.getByName('EAN-13') ;
            }
    catch(ex){
    var stL = docRef.layers.add();
            stL.name = "EAN-13";
                    }        
    EANGroup.move(stL, ElementPlacement.PLACEATEND);
    
    }



function replace_string(txt,cut_str,paste_str){ 
var f=0;
var ht='';
ht = ht + txt;
f=ht.indexOf(cut_str);
while (f!=-1){ 
//цикл для вырезания всех имеющихся подстрок 
f=ht.indexOf(cut_str);
if (f>0){
ht = ht.substr(0,f) + paste_str + ht.substr(f+cut_str.length);
};
};
return ht
};



function totext(){
    
    var over12 = dlg.alertBtnsPnl2.titleEt.text;
    if (over12.length >12 )
    dlg.alertBtnsPnl2.titleEt.text = over12.substring(0,12);
    var chk13 = SUM13(over12);    
    dlg.alertBtnsPnl2.TirSt.text = chk13;    
    EAN = over12+chk13;
    
    }


 


function CreatEAN(){

zX = 5;
zY = 5;

var chkSum13=SUM13(EAN);

// Начинаем рисовать штрихкод

SE();                                                                // стартовый блок

zX+=blkE;                                                        // смещение от первого блока
numBlokA1();                                                    // первый цифровой блок. Он всегда тип А
        
switch    (EAN.charAt(0)){

        case '0':
        for (var j=2;j<7;j++){
                numBlokAB(tablEAN[0].charAt(j-1),j); //  в зависимости от первой цифры кода выбираем последовательность АВ блоков из таблицы
                zX+=blkD;
                }
                CENTER();                                       // центральный блок
                zX+=blkC; 
        for (var u=7;u<13;u++){
                numBlokC(u);                                    // правая часть штрихкода - блок С
                zX+=blkD;
            }

        break;
        case '1':
        for (var j=2;j<7;j++){
                numBlokAB(tablEAN[1].charAt(j-1),j);
                zX+=blkD;
                }
                CENTER();
                zX+=blkC; 
        for (var u=7;u<13;u++){
                numBlokC(u);
                zX+=blkD;
            }

        break;
        case '2':
        for (var j=2;j<7;j++){
                numBlokAB(tablEAN[2].charAt(j-1),j);
                zX+=blkD;
                }
                CENTER();
                zX+=blkC; 
        for (var u=7;u<13;u++){
                numBlokC(u);
                zX+=blkD;
            }

        break;
        case '3':
        for (var j=2;j<7;j++){
                numBlokAB(tablEAN[3].charAt(j-1),j);
                zX+=blkD;
                }
                CENTER();
                zX+=blkC; 
        for (var u=7;u<13;u++){
                numBlokC(u);
                zX+=blkD;
            }

        break;
        case '4':
        for (var j=2;j<7;j++){
                numBlokAB(tablEAN[4].charAt(j-1),j);
                zX+=blkD;
                }
                CENTER();
                zX+=blkC; 
        for (var u=7;u<13;u++){
                numBlokC(u);
                zX+=blkD;
            }

        break;
        case '5':
        for (var j=2;j<7;j++){
                numBlokAB(tablEAN[5].charAt(j-1),j);
                zX+=blkD;
                }
                CENTER();
                zX+=blkC; 
        for (var u=7;u<13;u++){
                numBlokC(u);
                zX+=blkD;
            }

        break;
        case '6':
        for (var j=2;j<7;j++){
                numBlokAB(tablEAN[6].charAt(j-1),j);
                zX+=blkD;
                }
                CENTER();
                zX+=blkC; 
        for (var u=7;u<13;u++){
                numBlokC(u);
                zX+=blkD;
            }

        break;
        case '7':
        for (var j=2;j<7;j++){
                numBlokAB(tablEAN[7].charAt(j-1),j);
                zX+=blkD;
                }
                CENTER();
                zX+=blkC; 
        for (var u=7;u<13;u++){
                numBlokC(u);
                zX+=blkD;
            }

        break;
        case '8':
        for (var j=2;j<7;j++){
                numBlokAB(tablEAN[8].charAt(j-1),j);
                zX+=blkD;
                }
                CENTER();
                zX+=blkC; 
        for (var u=7;u<13;u++){
                numBlokC(u);
                zX+=blkD;
            }

        break;
        case '9':
        for (var j=2;j<7;j++){
                numBlokAB(tablEAN[9].charAt(j-1),j);
                zX+=blkD;
                }
                CENTER();
                zX+=blkC; 
        for (var u=7;u<13;u++){
                numBlokC(u);
                zX+=blkD;
            }
                
        break;

    }
                SE();           // конечный блок    
   
textEAN(); // Create digit TEXT for barcode
EANGroup.position =Array (XGr*mm,-YGr*mm); // Move  group barcode to position 
redraw();

}

//============== Function create text number code
function textEAN(){



zX = 5;
zY = 5;    
var pointTextRef1 = EANGroup.textFrames.add();
pointTextRef1.textRange.size = 9;
pointTextRef1.contents = EAN.charAt(0);
pointTextRef1.top = (zY-blkH)*mm;
pointTextRef1.left = (zX-2)*mm;
pointTextRef1.textRange.characterAttributes.textFont =  textFonts.getByName("ocrb10");

var pointTextRef2 = EANGroup.textFrames.add();
pointTextRef2.textRange.size = 9;
pointTextRef2.contents = EAN.substring(1,7);
pointTextRef2.top = (zY-blkH)*mm;
pointTextRef2.left = (zX+1)*mm;
pointTextRef2.textRange.characterAttributes.textFont =  textFonts.getByName("ocrb10");

var pointTextRef3 = EANGroup.textFrames.add();
pointTextRef3.textRange.size = 9;
pointTextRef3.contents = EAN.substring(7,13);
pointTextRef3.top = (zY-blkH)*mm;
pointTextRef3.left = (zX+16)*mm;
pointTextRef3.textRange.characterAttributes.textFont =  textFonts.getByName("ocrb10");


    
    }

//============ Функция отрисовки первого блока левой части. Он всегда типа А
function numBlokA1(){

    switch (EAN.charAt(1)){
                    case '0':
                        A_0();
                    break;            
                    case '1':
                        A_1();
                    break;            
                    case '2':
                        A_2();
                    break;            
                    case '3':
                        A_3();
                    break;            
                    case '4':
                        A_4();
                    break;            
                    case '5':
                        A_5();
                    break;            
                    case '6':
                        A_6();
                    break;            
                    case '7':
                        A_7();
                    break;            
                    case '8':
                        A_8();
                    break;            
                    case '9':
                        A_9();
                    break;            
            
            }
zX+=blkD;
    }


//============ Функция отрисовки правой части штрихкода. Он всегда типа С
function numBlokC(numC){

    switch (EAN.charAt(numC)){
        case '0':
        C_0();
        break;
        case '1':
        C_1();
        break;
        case '2':
        C_2();
        break;
        case '3':
        C_3();
        break;
        case '4':
        C_4();
        break;
        case '5':
        C_5();
        break;
        case '6':
        C_6();
        break;
        case '7':
        C_7();
        break;
        case '8':
        C_8();
        break;
        case '9':
        C_9();
        break;
        }

}

//============ Функция отрисовки блока левой части начиная со второй позиции.  В зависимости от таблицы numBlokAB.
function numBlokAB(ab,digBlok) {
    
    switch (ab){
        case 'A':
       switch (EAN.charAt(digBlok)){
                    case '0':
                        A_0();
                    break;            
                    case '1':
                        A_1();
                    break;            
                    case '2':
                        A_2();
                    break;            
                    case '3':
                        A_3();
                    break;            
                    case '4':
                        A_4();
                    break;            
                    case '5':
                        A_5();
                    break;            
                    case '6':
                        A_6();
                    break;            
                    case '7':
                        A_7();
                    break;            
                    case '8':
                        A_8();
                    break;            
                    case '9':
                        A_9();
                    break;            
            
                                    }
                    break;
                  
        case 'B':
   switch (EAN.charAt(digBlok)){
                    case '0':
                        B_0();
                    break;            
                    case '1':
                        B_1();
                    break;            
                    case '2':
                        B_2();
                    break;            
                    case '3':
                        B_3();
                    break;            
                    case '4':
                        B_4();
                    break;            
                    case '5':
                        B_5();
                    break;            
                    case '6':
                        B_6();
                    break;            
                    case '7':
                        B_7();
                    break;            
                    case '8':
                        B_8();
                    break;            
                    case '9':
                        B_9();
                    break;            
            
            }    
            break;
    
    
    }                
    }

// расчет контрольного числа - 13 цифры.
function SUM13(EAN12){
var sumSt1;
var sumSt2;
if (EAN12.length < 12)
sumSt2 ="<?>";
else {

sumSt1 =  parseInt(EAN12.charAt(1))+parseInt(EAN12.charAt(3))+parseInt(EAN12.charAt(5))+parseInt(EAN12.charAt(7))+parseInt(EAN12.charAt(9))+parseInt(EAN12.charAt(11));
sumSt1 *=3;
sumSt1 += parseInt(EAN12.charAt(0))+parseInt(EAN12.charAt(2))+parseInt(EAN12.charAt(4))+parseInt(EAN12.charAt(6))+parseInt(EAN12.charAt(8))+parseInt(EAN12.charAt(10));
sumSt2 = sumSt1%10;
 if (!(sumSt2 == 0))
                {
                    sumSt2 = 10 - sumSt2;
                }
			
else {
	sumSt2 = 0 ;
	
	}			
    }
return sumSt2;
    }


// функция отрисовки прямоугольника (левый угол X, левый угол Y, ширина, высота, делать ли прямоугольник guideline, залочить прямоугольник) с возможностью  сделать его  в виде guideline
function rectGuide(y1,x1,RGw,RGh,gd,lock) {
	var rect = EANGroup.pathItems.rectangle( x1*mm, y1*mm, RGw*mm, RGh*mm );
	rect.stroked = true;
	rect.filled = false;
	rect.guides = gd; // это св-во как раз и делает направляющие из линии
	rect.locked = lock; //заблокироваnm направляющие, 
}

//функция отрисовки прямоугольника (левый угол X, левый угол Y, ширина, высота, цвет заливки) без обводки

function rect(y1,x1,Rw,Rh,colorFill) {
	var rect = EANGroup.pathItems.rectangle( x1*mm, y1*mm, Rw*mm, Rh*mm );
      
	rect.stroked = false;
	rect.filled = true;
    rect.fillColor = colorFill;
}


// Отрисовка  блоков тип A, B, C

function A_0(){
  rect(zX+blk*3,zY,blk*2,blkH,barColor);   
  rect(zX+blk*6,zY,blk,blkH,barColor);   
    }
function A_1(){
  rect(zX+blk*2,zY,blk*2,blkH,barColor);   
  rect(zX+blk*6,zY,blk,blkH,barColor);   
    }
function A_2(){
  rect(zX+blk*2,zY,blk,blkH,barColor);   
  rect(zX+blk*5,zY,blk*2,blkH,barColor);   
    }
function A_3(){
  rect(zX+blk,zY,blk*4,blkH,barColor);   
  rect(zX+blk*6,zY,blk,blkH,barColor);   
    }
function A_4(){
  rect(zX+blk,zY,blk,blkH,barColor);   
  rect(zX+blk*5,zY,blk*2,blkH,barColor);   
    }
function A_5(){
  rect(zX+blk,zY,blk*2,blkH,barColor);   
  rect(zX+blk*6,zY,blk,blkH,barColor);   
    }
function A_6(){
  rect(zX+blk,zY,blk,blkH,barColor);   
  rect(zX+blk*3,zY,blk*4,blkH,barColor);   
    }
function A_7(){
  rect(zX+blk,zY,blk*3,blkH,barColor);   
  rect(zX+blk*5,zY,blk*2,blkH,barColor);   
    }
function A_8(){
  rect(zX+blk,zY,blk*2,blkH,barColor);   
  rect(zX+blk*4,zY,blk*3,blkH,barColor);   
    }
function A_9(){
  rect(zX+blk*3,zY,blk,blkH,barColor);   
  rect(zX+blk*5,zY,blk*2,blkH,barColor);   
    }

function B_0(){
  rect(zX+blk,zY,blk,blkH,barColor);   
  rect(zX+blk*4,zY,blk*3,blkH,barColor);   
    }
function B_1(){
  rect(zX+blk,zY,blk*2,blkH,barColor);   
  rect(zX+blk*5,zY,blk*2,blkH,barColor);   
    }
function B_2(){
  rect(zX+blk*2,zY,blk*2,blkH,barColor);   
  rect(zX+blk*5,zY,blk*2,blkH,barColor);   
    }
function B_3(){
  rect(zX+blk,zY,blk,blkH,barColor);   
  rect(zX+blk*6,zY,blk,blkH,barColor);   
    }
function B_4(){
  rect(zX+blk*2,zY,blk*3,blkH,barColor);   
  rect(zX+blk*6,zY,blk,blkH,barColor);   
    }
function B_5(){
  rect(zX+blk,zY,blk*3,blkH,barColor);   
  rect(zX+blk*6,zY,blk,blkH,barColor);   
    }
function B_6(){
  rect(zX+blk*4,zY,blk,blkH,barColor);   
  rect(zX+blk*6,zY,blk,blkH,barColor);   
    }
function B_7(){
  rect(zX+blk*2,zY,blk,blkH,barColor);   
  rect(zX+blk*6,zY,blk,blkH,barColor);   
    }
function B_8(){
  rect(zX+blk*3,zY,blk,blkH,barColor);   
  rect(zX+blk*6,zY,blk,blkH,barColor);   
    }
function B_9(){
  rect(zX+blk*2,zY,blk,blkH,barColor);   
  rect(zX+blk*4,zY,blk*3,blkH,barColor);   
    }

function C_0(){
  rect(zX,zY,blk*3,blkH,barColor);   
  rect(zX+blk*5,zY,blk,blkH,barColor);   
    }
function C_1(){
  rect(zX,zY,blk*2,blkH,barColor);   
  rect(zX+blk*4,zY,blk*2,blkH,barColor);   
    }
function C_2(){
  rect(zX,zY,blk*2,blkH,barColor);   
  rect(zX+blk*3,zY,blk*2,blkH,barColor);   
    }
function C_3(){
  rect(zX,zY,blk,blkH,barColor);   
  rect(zX+blk*5,zY,blk,blkH,barColor);   
    }
function C_4(){
  rect(zX,zY,blk,blkH,barColor);   
  rect(zX+blk*2,zY,blk*3,blkH,barColor);   
    }
function C_5(){
  rect(zX,zY,blk,blkH,barColor);   
  rect(zX+blk*3,zY,blk*3,blkH,barColor);   
    }
function C_6(){
  rect(zX,zY,blk,blkH,barColor);   
  rect(zX+blk*2,zY,blk,blkH,barColor);   
    }
function C_7(){
  rect(zX,zY,blk,blkH,barColor);   
  rect(zX+blk*4,zY,blk,blkH,barColor);   
    }
function C_8(){
  rect(zX,zY,blk,blkH,barColor);   
  rect(zX+blk*3,zY,blk,blkH,barColor);   
    }
function C_9(){
  rect(zX,zY,blk*3,blkH,barColor);   
  rect(zX+blk*4,zY,blk,blkH,barColor);   
    }

// Отрисовка  блоков типа Начало и Конец
function SE(){
    
  rect(zX,zY,blk,blkHE,barColor);   
  rect(zX+blk*2,zY,blk,blkHE,barColor);  

    }

// Отрисовка  блоков тип в Центре
function CENTER(){
  rect(zX+blk,zY,blk,blkHE,barColor);   
  rect(zX+blk*3,zY,blk,blkHE,barColor);   
    }
function writeINI(){

           var openF = myINIFile.open('e'); // 'e' read-write open file.
          
myINIFile. writeln(win.tPanel0.tTab0.coord.name1.e1.text);
myINIFile. writeln(win.tPanel0.tTab0.coord.name1.e2.text);
if (win.tPanel0.tTab1.dopparam.name3.chkLayer.value == true)
myINIFile. writeln('Y');
else
myINIFile. writeln('N');
if (win.tPanel0.tTab1.dopparam.name3.chkScale.value == true){
myINIFile. writeln('Y');
myINIFile. writeln(win.tPanel0.tTab1.dopparam.name3.eScale.text);
}
else{
myINIFile. writeln('N');
myINIFile. writeln('100');
}

myINIFile. writeln('Y');
myINIFile. writeln(win.tPanel0.tTab1.pages.ePages.text);
myINIFile. writeln(win.tPanel0.tTab1.name4.eColumn.text);
myINIFile. writeln(win.tPanel0.tTab1.name4.eRow.text);
myINIFile. writeln(win.tPanel0.tTab1.name5.eDistanceX.text);
myINIFile. writeln(win.tPanel0.tTab1.name5.eDistanceY.text);
                                                                             
myINIFile. writeln(destFolder);
myINIFile.close();          
}

