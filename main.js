const fs = require('fs')
const remote = require('electron')
const clipboard = remote.clipboard
const excelToJson = require('convert-excel-to-json');

const name = "SUIVI_DES_PJ.xlsx"
const table = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]
const ALLTABLE = ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V"]


function writeAll(result) {
    document.body.innerHTML = ""
    document.body.style = ""
    let List = document.createElement("div")
    List.setAttribute("id", "List")
    document.body.appendChild(List)
    let tabl = document.createElement("span")
    tabl.innerHTML = table
    tabl.style.position = "fixed"
    tabl.style.top = "25px"
    tabl.style.left = "0px"
    document.body.appendChild(tabl)

    let b1 = document.createElement("button")
    b1.innerHTML = "Montrer tout transmis"
    b1.addEventListener("click", () => {
        let doc = document.getElementsByClassName('all')
        for (const d of doc) {d.style.display = 'none'}
        doc = document.getElementsByClassName('Transmis')
        for (const d of doc) {d.style.display = 'block'}
    })
    let b2 = document.createElement("button")
    b2.innerHTML = "Montrer Fini"
    b2.addEventListener("click", () => {
        let doc = document.getElementsByClassName('all')
        for (const d of doc) {d.style.display = "none"}
        doc = document.getElementsByClassName('Fini')
        for (const d of doc) {d.style.display = "block"}
    })
    let b3 = document.createElement("button")
    b3.innerHTML = "Show all"
    b3.addEventListener("click", () => {
        let doc = document.getElementsByClassName('all')
        for (const d of doc) { d.style.display = "block"}
    })
    let b4 = document.createElement("button")
    b4.innerHTML = "reset"
    b4.addEventListener("click", () => {
        writeAll(result)
    })
    let input = document.createElement("input")
    let b5 = document.createElement("button")
    b5.innerHTML = "Zone Select"
    b5.addEventListener("click", () => {
        let size = input.value
        let all = document.getElementsByClassName('all')
        let sizeSplited = size.split(" ")
        sizeSplited.splice(0, 1)
        for (d of all) {d.style.display = "none"}
        for (let i = size.charCodeAt(0); i <= size.charCodeAt(1); i++) {
            let letter = String.fromCharCode(i)
            let doc = document.getElementsByClassName("name"+letter)
            for (const d of doc) {
                let mid = d.className.split(" ")[1]
                for (r of sizeSplited) {if (mid == r) { d.style.display = "block" }}
            }
        }
    })
    let searchN = document.createElement("input")
    let b6 = document.createElement("button")
    b6.innerHTML = "Tout Select"
    b6.addEventListener("click", () => {
        let size = searchN.value
        let all = document.getElementsByClassName('all')
        let sizeSplited = size.split(" ")
        let num = sizeSplited[0].split(",")
        sizeSplited.splice(0, 1)
        for (d of all) {
            d.style.display = "none"
            for (const n of num) {
                let a = n.split("&")
                let value = true
                for (const b of a) {
                    if (!d.innerText.includes(b.toString())) value = false
                }
                if (value) {
                    let mid = d.className.split(" ")[1]
                    for (r of sizeSplited) {if (mid == r) { d.style.display = "block" }}
                }
            }
        }
    })
    let b7 = document.createElement("button")
    b7.innerHTML = "Montrer Refus"
    b7.addEventListener("click", () => {
        let doc = document.getElementsByClassName('all')
        for (const d of doc) {d.style.display = "none"}
        doc = document.getElementsByClassName('Refus')
        for (const d of doc) {d.style.display = "block"}
    })
    let searchS = document.createElement("input")
    let b8 = document.createElement("button")
    b8.innerHTML = 'Select Spe'
    b8.addEventListener("click", () => {
        let size = searchS.value
        let all = document.getElementsByClassName('all')
        let sizeSplited = size.split(" ")
        let num = sizeSplited[0].split(",")
        sizeSplited.splice(0, 1)
        for (d of all) {
            d.style.display = "none"
            for (const n of num) {
                let a = n.split("&")
                let nv = false
                for (const b of a) {
                    if (b.charAt(1) == ':') {
                        let c = b.split(":")
                        let r = c[1].split("/")
                        for (let i = r[0]; i <= r[1]; i++) if (!d.getElementsByClassName(c[0])[0].innerText.includes(i.toString())) nv = true; else {nv = false; break}
                    } else if (b.charAt(1) == '!') {
                        let c = b.split("!")
                        if (!d.getElementsByClassName(c[0])[0].innerText.includes(c[1])) { nv = true; break }
                    }
                }
                if (!nv) {
                    let mid = d.className.split(" ")[1]
                    for (r of sizeSplited) {if (mid == r) { d.style.display = "block" }}
                }
            }
        }
    })
    b1.setAttribute("id", "b1")
    b2.setAttribute("id", "b2")
    b3.setAttribute("id", "b3")
    b4.setAttribute("id", "b4")
    b5.setAttribute("id", "b5")
    b6.setAttribute("id", "b6")
    b7.setAttribute("id", "b7")
    b8.setAttribute("id", "b8")
    input.setAttribute("id", "inputName")
    searchN.setAttribute("id", "inputNumber")
    searchS.setAttribute("id", "inputSpan")
    List.appendChild(b1)
    List.appendChild(b2)
    List.appendChild(b3)
    List.appendChild(b4)
    List.appendChild(b5)
    List.appendChild(b6)
    List.appendChild(b7)
    List.appendChild(b8)
    List.appendChild(input)
    List.appendChild(searchN)
    List.appendChild(searchS)
    for (const row of result) {
        let Etudiant
        if (document.getElementById(row["G"].toString()) == null) {
            Etudiant = document.createElement("div")
            Etudiant.setAttribute("class", "all")
            Etudiant.setAttribute("id", row["G"])
            let annee = document.createElement("span")
            annee.innerHTML = row["C"]
            annee.setAttribute("class", "A")
            let num = document.createElement("span")
            num.innerHTML = row["G"]
            num.setAttribute("class", "B")
            let nom = document.createElement("span")
            nom.innerHTML = row["K"]
            nom.setAttribute("class", "C")
            let prenom = document.createElement("span")
            prenom.innerHTML = row["L"]
            prenom.setAttribute("class", "D")
            let nbPJ = document.createElement("span")
            nbPJ.innerHTML = (row["O"] == "O") ? 1 : 0
            nbPJ.setAttribute("class", "E")
            let vali = document.createElement("span")
            vali.innerHTML = (row["P"] == "V") ? 1 : 0
            vali.setAttribute("class", "F")
            let trans = document.createElement("span")
            trans.innerHTML = (row["P"] == "T") ? 1 : 0
            trans.setAttribute("class", "G")
            let atten = document.createElement("span")
            atten.innerHTML = (row["P"] == "A") ? 1 : 0
            atten.setAttribute("class", "H")
            let refu = document.createElement("span")
            refu.innerHTML = (row["P"] == "R") ? 1 : 0
            refu.setAttribute("class", "I")
            let firstnum = document.createElement("span")
            firstnum.innerHTML = row["G"].toString().substring(0, 3)
            firstnum.setAttribute("class", "J")
            let comment = document.createElement("span")
            comment.innerHTML = (row["M"] == "RECAW") ? row["V"] : ""
            comment.setAttribute("class", "K")
            Etudiant.appendChild(annee)
            Etudiant.appendChild(num)
            Etudiant.appendChild(nom)
            Etudiant.appendChild(prenom)
            Etudiant.appendChild(nbPJ)
            Etudiant.appendChild(vali)
            Etudiant.appendChild(trans)
            Etudiant.appendChild(atten)
            Etudiant.appendChild(refu)
            Etudiant.appendChild(firstnum)
            Etudiant.appendChild(comment)

            Etudiant.addEventListener("auxclick", () => {
                clipboard.writeText(Etudiant.id, "selection");
                for(const cla of document.getElementsByClassName("all")) {cla.style.backgroundColor = ""; cla.style.color = ""}
                Etudiant.style.backgroundColor= "yellow"
                Etudiant.style.color = "black"
            })
            List.appendChild(Etudiant)
        } else {
            Etudiant = document.getElementById(row["G"].toString())
            let total = Etudiant.getElementsByClassName('E')[0]
            if (row["O"] == "O") {
                total.innerHTML = parseInt(total.innerHTML) + 1
                switch(row["P"]) {
                    case "V":
                        Etudiant.getElementsByClassName('F')[0].innerHTML = parseInt(Etudiant.getElementsByClassName('F')[0].innerHTML) + 1
                        break
                    case "T":
                        Etudiant.getElementsByClassName('G')[0].innerHTML = parseInt(Etudiant.getElementsByClassName('G')[0].innerHTML) + 1
                        break
                    case "A":
                        Etudiant.getElementsByClassName('H')[0].innerHTML = parseInt(Etudiant.getElementsByClassName('H')[0].innerHTML) + 1
                        break
                    case "R":
                        Etudiant.getElementsByClassName('I')[0].innerHTML = parseInt(Etudiant.getElementsByClassName('I')[0].innerHTML) + 1
                        break
                }
                if (row["M"] == "RECAW") Etudiant.getElementsByClassName('K')[0].innerHTML = row["V"]
            }
            let trans = Etudiant.getElementsByClassName('G')[0]
            let vali = Etudiant.getElementsByClassName('F')[0]
            let att = Etudiant.getElementsByClassName('H')[0]
            let oneAtt = (parseInt(att.innerHTML) == 1) ? true : false
            let d = new Date()
            let isyear = d.getUTCFullYear().toString().slice(2) == row["G"].toString().substring(1, 3)
            let clas = Etudiant.getAttribute("class")
            if (vali.innerHTML == total.innerHTML) {
                Etudiant.setAttribute("class", "all Fini name"+ row["K"].charAt(0)+ " anne" + row["G"].toString().substring(0, 3))
            } else if ((parseInt(trans.innerHTML) + parseInt(vali.innerHTML) == total.innerHTML) || (parseInt(trans.innerHTML) + parseInt(vali.innerHTML) == ((oneAtt) ? parseInt(total.innerHTML) - parseInt(att.innerHTML) : parseInt(total.innerHTML))) && isyear ) {
                Etudiant.setAttribute("class", "all Transmis name"+ row["K"].charAt(0)+ " anne" + row["G"].toString().substring(0, 3))
                if (oneAtt) { Etudiant.setAttribute("class", "all Transmis name"+ row["K"].charAt(0)+ " anne" + row["G"].toString().substring(0, 3) + " TA") }
            } else {
                let ref = Etudiant.getElementsByClassName('I')[0]
                if (ref.innerHTML > 0) Etudiant.setAttribute("class", "all Refus name"+ row["K"].charAt(0)+ " anne" + row["G"].toString().substring(0, 3));
                else Etudiant.setAttribute("class", "all Attente name"+ row["K"].charAt(0)+ " anne" + row["G"].toString().substring(0, 3));
            }
        }
    }
}

async function onclickedbutton() {
    const result = excelToJson({
        sourceFile: "../" + name,
        header: {rows:2, colums:3}
    })[" liste inscrits"]
    for (const row of result ) {
        for (const letter of ALLTABLE ) {
            if (row[letter] == undefined) { row[letter] = "" }
        }
    }
    writeAll(result)

}
