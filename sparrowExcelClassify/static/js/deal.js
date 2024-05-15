function chick(e) {
    for (let i = 0; i < sheet_names.length; i++) {
        let sheet_name = sheet_names[i]
        let choose = document.getElementById(sheet_name + "_choose")
        if (choose.checked === true) {
            let start = document.getElementById(sheet_name + "_start")
            let classify = document.getElementById(sheet_name + "_classify")
            if (start.value === "") {
                e.preventDefault();
                alert(sheet_name + "开始行列不能为空")
                break
            }
            if (end.value === "") {
                e.preventDefault();
                alert(sheet_name + "结束行列不能为空")
                break
            }
            if (classify.value === "") {
                e.preventDefault();
                alert(sheet_name + "分拣列不能为空")
                break
            }
        }
    }
}


function setData() {
    let root = document.getElementById("root")
    let sheets = document.createElement("input")
    sheets.type = "hidden"
    sheets.name = "sheets"
    sheets.value = sheet_names
    root.appendChild(sheets)
    for (let i = 0; i < sheet_names.length; i++) {
        let sheet_name = sheet_names[i]
        let div = document.createElement('div')
        div.id = sheet_name
        let h3 = document.createElement("h3")
        h3.innerText = sheet_name
        let input_h3 = document.createElement("input")
        input_h3.type = "hidden"
        input_h3.id = sheet_name + "_sheet_name"
        input_h3.name = sheet_name + "_sheet_name"
        input_h3.value = sheet_name
        let start_label = document.createElement('label')
        start_label.htmlFor = sheet_name + "_start"
        start_label.innerText = "开始行,如1:"
        let start_input = document.createElement('input')
        start_input.id = sheet_name + "_start"
        start_input.name = sheet_name + "_start"
        start_input.type = "number"
        let classify_label = document.createElement('label')
        classify_label.htmlFor = sheet_name + "_classify"
        classify_label.innerText = "根据那一列进行分拣:"
        let classify_input = document.createElement('input')
        classify_input.id = sheet_name + "_classify"
        classify_input.name = sheet_name + "_classify"
        classify_input.type = "number"
        let choose_label = document.createElement('label')
        choose_label.htmlFor = sheet_name + "_choose"
        choose_label.innerText = "此表需要被分拣"
        let choose_input = document.createElement('input')
        choose_input.id = sheet_name + "_choose"
        choose_input.name = sheet_name + "_choose"
        choose_input.value = "yes"
        choose_input.type = "radio"
        let not_choose_label = document.createElement('label')
        not_choose_label.htmlFor = sheet_name + "_not_choose"
        not_choose_label.innerText = "此表不需要被分拣,默认每个工作铺都会有相同的此表"
        let not_choose_input = document.createElement('input')
        not_choose_input.id = sheet_name + "_not_choose"
        not_choose_input.name = sheet_name + "_choose"
        not_choose_input.value = "no"
        not_choose_input.type = "radio"
        not_choose_input.checked = true
        let br1 = document.createElement('br')
        let br2 = document.createElement('br')
        let br3 = document.createElement('br')
        div.appendChild(h3)
        div.appendChild(input_h3)
        div.appendChild(start_label)
        div.appendChild(start_input)
        div.appendChild(br1)
        div.appendChild(classify_label)
        div.appendChild(classify_input)
        div.appendChild(br2)
        div.appendChild(choose_label)
        div.appendChild(choose_input)
        div.appendChild(br3)
        div.appendChild(not_choose_label)
        div.appendChild(not_choose_input)
        root.appendChild(div)
    }
}