const XLSX = require('xlsx');
const fs = require('fs');

class WCU {
    static Lni2Obj(source) {
        let text = fs.readFileSync(source).toString()
        text = text.replace(/^--.*$\r\n/gm, '') // s1
        text = text.replace(/^([_a-zA-Z0-9]+) = (.+)$/gm, '"$1": $2') // s2
        text = text.replace(/^\"([_a-zA-Z0-9]+)\"\: \{([^\}]+)\}$/gm, '"$1": [$2]') // s3
        text = text.replace(/\[([\r\n]+("\d+": .+[\r\n]+)*)\]/gm, '{$1}') // s4
        text = text.replace(/^(\d+)\: (.+)/gm, '"$1": $2') // s6
        text = text.slice(0, text.length - 1) // remove head & foot "{" "}" added by "s4"
        // console.log(text)
        let objs = text.split('\r\n\r\n')
        let arrs = []
        objs.forEach((obj, index, array) => {
            // console.log(index, obj)
            let opt = obj.trim().replace(/^\[([a-zA-Z0-9]+)\]\r\n([\s\S]+)/gm, function (source, a, b) {
                let list = b.split('\r\n')
                // console.log(a)
                // console.log(`source\n${source}\na\n${a}\nb\n${b}`)
                for (let i = 0; i < list.length; i++) {
                    let line = list[i]
                    let last = line[line.length - 1]
                    if (last != '[' && last != '{') {
                        // console.log(`line ${i}: ${line}`)
                    }
                    // list[i] = line
                }
                b = list.join(',\r\n')
                b = b.replace(/,,/gm, ',').replace(/\[,/gm, '[').replace(/\{,/gm, '{').replace(/,\r\n\}/gm, '\r\n}').replace(/,\r\n\]/gm, '\r\n]')
                return `"${a}": {\r\n${b}\r\n}`
            })
            // opt = opt.slice(0, opt.length - 1)
            // console.log(opt)
            arrs.push(opt)
        })
        text = arrs.join(',\r\n')
        text = `{\r\n${text}\r\n}`
    
        // fs.writeFileSync("d:/ability.ini", text)
    
        let json = JSON.parse(text)
    
        return json
    }
    static Obj2Lni(json, target) {
        let result = []
        for (const id in json) {
            if (json.hasOwnProperty(id)) {
                const obj = json[id]
                result.push(`[${id}]`)
                for (const field in obj) {
                    if (obj.hasOwnProperty(field)) {
                        const value = obj[field]
                        if (value.constructor == Array) {
                            result.push(`${field} = {`)
                            value.forEach(item => {
                                if (item.constructor == String) {
                                    result.push(`"${item}",`)
                                }
                                else if (item.constructor == Number) {
                                    result.push(`${item},`)
                                }
                            })
                            result.push(`}`)
                        }
                        else if (value.constructor == Object) {
                            result.push(`${field} = {`)
                            for (const itemId in value) {
                                if (value.hasOwnProperty(itemId)) {
                                    const item = value[itemId]
                                    if (item.constructor == String) {
                                        result.push(`${itemId} = "${JSON.stringify(item)}",`)
                                    }
                                    else if (item.constructor == Number) {
                                        result.push(`${itemId} = ${item},`)
                                    }
                                }
                            }
                            result.push(`}`)
                        }
                        else if (value.constructor == String) {
                            result.push(`${field} = ${JSON.stringify(value)}`)
                        }
                        else if (value.constructor == Number) {
                            result.push(`${field} = ${value}`)
                        }
                    }
                }
                result.push(``)
            }
        }
        let output = result.join('\r\n')
        fs.writeFileSync(target, output)
    }
    static Json2Obj(source) {
        let text = fs.readFileSync(source).toString()
        let json = JSON.parse(text)
        return json
    }
    static Obj2Json(json, target) {
        let output = JSON.stringify(json)
        fs.writeFileSync(target, output)
    }
    static Lni2Json(source, target) {
        var json = WCU.Lni2Obj(source)
    
        text = JSON.stringify(json, null, 2)
        fs.writeFileSync(target, text)
    
        // fs.writeFileSync(target, `${text}`)
    }
    static Json2Lni(source, target) {
        let json = WCU.Json2Obj(source)
        WCU.Obj2Lni(json, target)
    }
    static LoadTable(filename) {
        let wb = XLSX.readFile(filename)
        WCU.sheets = wb.Sheets
    }
    static GetSheet(sheetname) {
        return XLSX.utils.sheet_to_json(WCU.sheets[sheetname], { header: 1 })
    }
    /**
     * @param {string} sheetname compile target sheet name.
     * @param {Function<object, integer, array>} action action for each sheet.
     */
    static ForSheet(sheetname, action) {
        let sheet = WCU.GetSheet(sheetname)
        for (let i = 1; i <= sheet.length - 1; i++) {
            let row = sheet[i]
            action(row, i, sheet)
        }
    }
    static SetResourcesRoot(source, target) {
        WCU.resSource = source
        WCU.resSource = target
    }
}

module.exports = WCU;
