var mammoth = require("mammoth");
var fs = require("fs");

mammoth.convertToHtml({path: "D:\\examinationPortal\\questions.docx"})
    .then(function(result){
        var html = result.value; // The generated HTML
        fs.writeFile("output.html", html, function(err) {
            if(err) {
                return console.log(err);
            }
            console.log("The file was saved!");
        }); 
    })
    .done();