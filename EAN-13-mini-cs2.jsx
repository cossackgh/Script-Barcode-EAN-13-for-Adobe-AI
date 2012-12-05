// Create Barcode EAN-13mini v 1.0.9
// (с) 2010-2011. Dmitry Yagupov 
// www.za-vod.ru
// info@za-vod.ru


//#target illustrator
var fos =  Folder.fs;
//Get Version AI
var versionAI = parseInt(version.substr(0,2));
//alert(versionAI);
var EANdlgStr = '';
//Get fullPath to Script file
if (fos == 'Windows'){
var pathJSX=$.fileName;
pathJSX = pathJSX.substring(0,pathJSX.lastIndexOf("\\"));
pathJSX = replace_string(pathJSX,'\\','/');
pathJSX = '/'+replace_string(pathJSX,':','');
}
else{
var pathJSX=Folder.desktop;
	}
var docRef = activeDocument;
var EANGroup = docRef.groupItems.add(); 

var mm=2.834645669; //convert point to mm
var blk = 0.33;
var blkD = blk*7;
var blkE= blk*3;
var blkC= blk*5;
var blkH = 22.85;
var blkHE =24.5;
var zX=0;
var zY=0;
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

// Set Zero point ruler on Document

var hDoc = docRef.height;
var wDoc = docRef.width;

docRef.rulerOrigin = Array(0, hDoc);



// Set color values for the CMYK object
var barColor = new CMYKColor();
barColor.black = 100;
barColor.cyan = 0;
barColor.magenta = 0;
barColor.yellow = 0;


//*******************************************
// Create Dialog Window

var res =
"dialog { alignChildren: 'fill', text: 'EAN-13 mini', \
digit12: Panel { orientation: 'column', \
text: 'EAN-13', \
name2: Group { orientation: 'row', \
s: StaticText { text:'Enter 12 digit code:' }, \
e: EditText { characters: 12 } ,\
sh: StaticText { text:'<?>' }, \
} \
}, \
buttons: Group { orientation: 'row', alignment: 'center', \
okBtn: Button { text:'OK', properties:{name:'ok'} }, \
cancelBtn: Button { text:'Cancel', properties:{name:'cancel'} } \
} \
}";


if (versionAI > 12){
win = new Window (res); 



try {
var imgbar = new File(pathJSX+'/barcode.png');

// Add Picture Barcode to Dialog Window

win.add("image",[16,16,116,57],imgbar);
}catch(e) {
			}

// Check If enter only digit 0-9
win.digit12.name2.e.onChanging = function (){    
	var nowEnter = win.digit12.name2.e.text;
	var vPattern = /[^0-9]/;
	var noneD = /\D/g;
	var result = vPattern.test(nowEnter);

if (result == true)
{
	nowEnter = nowEnter.replace(noneD, "") ;
	win.digit12.name2.e.text = nowEnter;
    alert('Only numbers are permitted for this field.');
}

	
	if ( nowEnter.length > 12) {
		alert('You enter more 12 digit');
		nowEnter = nowEnter.substring(0,12);
		win.digit12.name2.e.text =  nowEnter;
		
		}
	
    var chk13 = SUM13(nowEnter);    

    EAN = nowEnter+chk13;	
	win.digit12.name2.sh.text = chk13;
    }

// OK botton Click
win.buttons.okBtn.onClick = function actionPlace() { 
    var enterDigits = win.digit12.name2.e.text.length;


    chkLayer();

            
    if ( enterDigits == 12){   
       
        CreatEAN(); 
        actionCanceled(); 
       
        }
    else 
    alert ('You NO Enter 12 digits');
    
}

win.buttons.cancelBtn.onClick = function exitDlg() { 
    win.close();
    }


win.center(); 
}
else win = "";


//CS-CS2

if (versionAI > 12){
win.show();
}
else {
var EANCS = prompt('Enter 12 first digit code', '') ;     
    if (EANCS != null){
    var enterDigits = EANCS.length;
    chkLayer();
     var chk13 = SUM13(EANCS);    
    EAN = EANCS+chk13;
    
   if ( enterDigits == 12){          
        CreatEAN();
        }
    else 
    alert ('You NO Enter 12 digits');
    }
    else alert ('You NO Enter nothing');
    
}

    

 function actionCanceled() { 
	win.close();
}


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



SE();                                                                

zX+=blkE;                                                      
numBlokA1();                                                   
        
switch    (EAN.charAt(0)){

        case '0':
        for (var j=2;j<7;j++){
                numBlokAB(tablEAN[0].charAt(j-1),j); 
                zX+=blkD;
                }
                CENTER();                                      
                zX+=blkC; 
        for (var u=7;u<13;u++){
                numBlokC(u);                                   
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
                SE();           
   
textEAN();
}

//============== Function create text number code
function textEAN(){

var posXGroup = 10; // X coordinate Barcode
var XGr = parseInt(posXGroup);    

var posYGroup = 10;  // Y coordinate Barcode
var YGr = parseInt(posYGroup);  

zX = 5;
zY = 5;    
var pointTextRef1 = EANGroup.textFrames.add();
pointTextRef1.textRange.size = 9;
pointTextRef1.contents = EAN.charAt(0);
pointTextRef1.top = (zY-23)*mm;
pointTextRef1.left = (zX-2)*mm;
pointTextRef1.textRange.characterAttributes.textFont =  textFonts.getByName("ocrb10");

var pointTextRef2 = EANGroup.textFrames.add();
pointTextRef2.textRange.size = 9;
pointTextRef2.contents = EAN.substring(1,7);
pointTextRef2.top = (zY-23)*mm;
pointTextRef2.left = (zX+1)*mm;
pointTextRef2.textRange.characterAttributes.textFont =  textFonts.getByName("ocrb10");

var pointTextRef3 = EANGroup.textFrames.add();
pointTextRef3.textRange.size = 9;
pointTextRef3.contents = EAN.substring(7,13);
pointTextRef3.top = (zY-23)*mm;
pointTextRef3.left = (zX+16)*mm;
pointTextRef3.textRange.characterAttributes.textFont =  textFonts.getByName("ocrb10");

EANGroup.position =Array (XGr*mm,-YGr*mm);

redraw();
    
    }

//============ 
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


//============ 
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

//============ 
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


function rectGuide(y1,x1,RGw,RGh,gd,lock) {
	var rect = EANGroup.pathItems.rectangle( x1*mm, y1*mm, RGw*mm, RGh*mm );
	rect.stroked = true;
	rect.filled = false;
	rect.guides = gd; 
	rect.locked = lock; 
}


function rect(y1,x1,Rw,Rh,colorFill) {
	var rect = EANGroup.pathItems.rectangle( x1*mm, y1*mm, Rw*mm, Rh*mm );
      
	rect.stroked = false;
	rect.filled = true;
    rect.fillColor = colorFill;
}



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


function SE(){
    
  rect(zX,zY,blk,blkHE,barColor);   
  rect(zX+blk*2,zY,blk,blkHE,barColor);  

    }


function CENTER(){
  rect(zX+blk,zY,blk,blkHE,barColor);   
  rect(zX+blk*3,zY,blk,blkHE,barColor);   
    }