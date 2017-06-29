const express = require('express')
const app = express()
var xlsx = require('node-xlsx');
var fs = require('fs');
var mailer = require("nodemailer");
var Mustache = require('mustache');
var Q = require('q');
// Use Smtp Protocol to send Email


app.get('/', function (req, res) {

    var file_contents = JSON.parse(fs.readFileSync(__dirname + '/setup.json'));
    console.log("SETUP CONTENTS>>", JSON.stringify(file_contents));
    if (file_contents.username && file_contents.password && file_contents.xlsx_file_name && file_contents.template_file_name && file_contents.data_index && file_contents.email_index) {
        try {
            console.log((__dirname + "/" + file_contents.xlsx_file_name));
            var obj = xlsx.parse(fs.readFileSync(__dirname + "/"+file_contents.xlsx_file_name));
        } catch (error) {
            res.json(error);
            console.log("ERROR IN READING XLSX FILE>>>", error);
        }
        var data = obj[0].data;
        console.log("data>>>>", JSON.stringify(data));
        transporter = mailer.createTransport({
            service: 'gmail',
            auth: {
                user: file_contents.username,
                pass: file_contents.password
            }
        });
        iterateArrayWithPromise(data, function (index, row) {
            if (index >= file_contents.data_index - 1 && row.length > 0) {
                return send_email(row, file_contents);
            }
        }).catch(function (error) {
            if (error.responseCode == 535) {
                console.log("Allow this Application to Send Mail using your gmail account");
            }
            console.log(error);
        });
        res.json({"code_name": "SUCCESS"});
    } else {
        res.json({"code_name": "ERROR", "message": "FILL ALL DETAILS IN SETUP FILE"});
    }
})

function send_email(row, file_contents) {
    var D = Q.defer();
    var fields = file_contents.fields;
    var render_data = {};
    for (var key in fields) {
        render_data[key] = row[fields[key] - 1]
    }

    var email = row[file_contents["email_index"] - 1];
    if (email) {

        var text = fs.readFileSync(__dirname + "/" + file_contents.template_file_name, "utf-8");
        text = Mustache.render(text, render_data);
        var mail = {
            from: file_contents.username,
            to: email,
            subject: Mustache.render(file_contents.email_subject, render_data),
            html: text
        }
        transporter.sendMail(mail, function (error, response) {
            if (error) {
                D.reject(error);
            } else {
                console.log("message sent");
                D.resolve()
            }
        });
    } else {
        D.reject();
    }
    return D.promise;
}

function iterateArrayWithPromise(array, task) {
    var D = Q.defer();
    var length = array ? array.length : 0;
    if (length == 0) {
        setTimeout(function () {
            D.resolve();
        }, 10)
        return D.promise;
    }
    var index = 0;

    function loop(index) {
        try {
            var onResolve = function () {
                index = index + 1;
                if (index == array.length) {
                    D.resolve();
                } else {
                    loop(index);
                }
            }
            try {
                var p = task(index, array[index]);
                if (!p) {
                    onResolve();
                    return;
                }
                p.then(onResolve)
                    .fail(function (err) {
                        D.reject(err);
                    })
            } catch (e) {
                D.reject(e);
            }
        } catch (e) {
            D.reject(e);
        }
    }

    loop(index);
    return D.promise;
}

app.listen(3000, function () {
    console.log('Example app listening on port 3000!')
})