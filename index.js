const express = require("express");
const Excel = require("exceljs")
const path = require("path");
const multer = require("multer");


const app = express();

app.set("views",path.join(__dirname,"views"));
app.set("view engine","ejs");

app.use(express.urlencoded({ extended: true }))
app.use(express.static("public"));

const classTime = {
	"2": "9:00-9:30",
	"3": "9:30-10:00",
	"4": "10:00-10:30",
	"5": "10:30-11:00",
	"6": "merged coloumn",
	"7": "11:00-11:30",
	"8": "11:30-12:00",
	"9": "12:00-12:30",
	"10": "12:30-13:00",
	"11": "13:00-13:30",
	"12": "13:30-14:00",
	"13": "merged coloumn",
	"14": "14:00-14:30",
	"15": "14:30-15:00",
	"16": "15:00-15:30",
	"17": "15:30-16:00",
	"18": "16:00-16:30",
	"19": "16:30-17:00",
	"20": "17:00-17:30",
	"21": "17:30-18:00",
	"22": "18:00-18:30",
	"23": "18:30-19:00",
	"24": "19:00-19:30",
	"25": "19:30-20:00",
}
const weekDay = {"0":"SUN","1":"MON","2":"TUE","3":"WED","4":"THU","5":"FRI","6":"SAT"};
 
const storage = multer.memoryStorage()
const upload = multer({ storage: storage })

app.get("/",(req,res) => {
    res.render("index");
})

app.post("/aipScheduling",upload.single("facultyRoutine"),async (req,res) => {
   var startTime = req.body.startTime;
   var endTime = req.body.endTime;
   var time = `${startTime}-${endTime}`;
   var date = req.body.date;
   var inputDate = new Date(date);
   var dayNumber = inputDate.getDay();
   var day = weekDay[dayNumber];
   console.log(day);
   var freeSlot = [];
   var filledSlot = [];
	if(req.file){
		if(req.file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'){
		const workBook = new Excel.Workbook();
		await workBook.xlsx.load(req.file.buffer).then(() =>{
			
			//This is a for each loop which will iterate over every sheet in the workbook.
			workBook.eachSheet(function(worksheet){
				
				for(let row = 7; row<= 12; row++) //This loop will control the row itretaion
				{
					let days = worksheet.getCell(row,1)
					// console.log(`WORKING DAY ${worksheet.getCell(row,1)}`);
					for(let col = 2; col <=25; col++) //This loop will control the coloumn iteration 
					{
						if(classTime[col] == time && days == day){
						//If the cell is having a value that means the facukty is not free else they are free and the meeting can be scedulecd for that time slot.
							if(!worksheet.getCell(row,col).value){
								let factName = worksheet.getCell(1,1).value
								freeSlot.push(`${factName.substring(29)}`);
								
							}
						}
						else if(day == "SUN"){
							var message = "There is no slot available for Sunday";
							res.render("error",{message});
						}
						// else if(classTime[col] != time){
						// 	var message = "Please enter time as per schedule";
						// 	res.render("error",{message});
						// }
						
					}
				}
			})
		})
		}
		else{
			var message = "File should be .xlsx file. Please try again with a valid file";
			res.render("error",{message});
		}
	}
	else{
		var message = "File missing. Choose the faculty shedule excel sheet. It should be in .xlsx format";
			res.render("error",{message});
	}
    
    res.render("aipSchedule",{freeSlot, filledSlot,day,time});
})

app.listen(3000,() => {
    console.log("Server Started On Port 3000");
})
