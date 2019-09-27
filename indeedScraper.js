const indeed = require('indeed-scraper');
var excel = require('excel4node');

var workBook = new excel.Workbook(); //Creates a new instance of the Workbook class

var headingStyle = workBook.createStyle({ //Creeates a styling object for the headers in the workbook
    font: {
        color: '#000000',
        size: 16,
        name: 'Times New Roman',
        bold: true,
    },
    alignment: {
        horizontal: ['center'],
        vertical: ['center'],
    },
    border: {
        left: {
            style:'thick',
            color:'#000000',
        },
        right: {
            style:'thick',
            color:'#000000',
        },
        top: {
            style:'thick',
            color:'#000000',
        },
        bottom: {
            style:'thick',
            color:'#000000',
        },
    },
});

var cellStyle = workBook.createStyle({
    font: {
        color: '#000000',
        size: 12,
        name: 'Times New Roman',
    },
    alignment: {
        horizontal: ['center'],
        vertical: ['center'],
        wrapText: true,
    },
});

var workSheet = workBook.addWorksheet('Jobs'); //Add worksheet

workSheet.cell(2,2).string('Title').style(headingStyle) //Selects target cell and inserts value and applies styling
workSheet.cell(2,3).string('Summary').style(headingStyle)
workSheet.cell(2,4).string('URL').style(headingStyle)
workSheet.cell(2,5).string('Company').style(headingStyle)
workSheet.cell(2,6).string('Location').style(headingStyle)
workSheet.cell(2,7).string('Post Date').style(headingStyle)
workSheet.cell(2,8).string('Salary').style(headingStyle)

var columnCount = 7;

for (var i = 2; i < 2 + columnCount; i++) { 
    workSheet.column(i).setWidth(30); //Loops through the columns in the workbook and sets their width
}

const queryOptions = {
    host: 'www.indeed.com',
    query: 'Software Engineering Intern',
    city: 'New York, NY',
    radius: '10',
    level: 'entry_level',
    jobType: 'full_time',
    maxAge: '14',
    sort: 'date',
    limit: 30
};

var rowCount = queryOptions.limit;

indeed.query(queryOptions).then(function (response) { //Sends an HTTP request to Indeed and returns with a Promise of objects inside of an array

    var jobKeys = Object.keys(response[0]); //Extracting the keys inside the response array

    for (var i = 3; i < 3 + rowCount; i++) { //Looping through the response array, and using the keys to extract the values from each object in the array
        for (var j = 2; j < 2 + columnCount; j++) {
            if (response[i-3].hasOwnProperty(jobKeys[j-2])) {
                workSheet.cell(i,j).string(response[i-3][jobKeys[j-2]]).style(cellStyle); //Inserting the values extracted into the excel worksheet
            }
        }
    }
    workBook.write('Jobs.xlsx'); //Finally, writing everything to a worksheet with a title of my choosing
})
