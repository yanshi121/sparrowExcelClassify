function handleSelectChange(event, name, index = 0) {
        let input = document.getElementById(name + "_value_" + index.toString())
        let label = document.getElementById(name + "_valueLabel_" + index.toString())
        let selectedValue = event.target.value;
        switch (selectedValue) {
            case "sum":
                input.type = "number"
                label.innerText = "所选择的列:"
                input.className = "input"
                input.style.display = "inline-block"
                label.style.display = "inline-block"
                input.name = name + "_sum_" + index.toString()
                break;
            case  "mean":
                input.type = "number"
                label.innerText = "所选择的列:"
                input.className = "input"
                input.style.display = "inline-block"
                label.style.display = "inline-block"
                input.name = name + "_mean_" + index.toString()
                break;
            case  "median":
                input.type = "number"
                label.innerText = "所选择的列:"
                input.className = "input"
                input.style.display = "inline-block"
                label.style.display = "inline-block"
                input.name = name + "_median_" + index.toString()
                break;
            case  "min":
                input.type = "number"
                label.innerText = "所选择的列:"
                input.className = "input"
                input.style.display = "inline-block"
                label.style.display = "inline-block"
                input.name = name + "_min_" + index.toString()
                break;
            case  "max":
                input.type = "number"
                label.innerText = "所选择的列:"
                input.className = "input"
                input.style.display = "inline-block"
                label.style.display = "inline-block"
                input.name = name + "_max_" + index.toString()
                break;
            case  "userDefined":
                input.style.display = "inline-block"
                input.type = "text"
                label.innerText = "输入自定义的值:"
                input.className = ""
                label.style.display = "inline-block"
                input.name = name + "_userDefined_" + index.toString()
                break;
            case  "sign":
                input.style.display = "none"
                label.style.display = "none"
                input.name = name + "_sign_" + index.toString()
                break;
            case  "num":
                input.style.display = "none"
                label.style.display = "none"
                input.name = name + "_num_" + index.toString()
                break;
            default:
                input.style.display = "none"
                label.style.display = "none"
        }
    }

function deleteNode(name, index = 0) {
        let parentNode = document.getElementById(name);
        let childNode = document.getElementById(name + "_div_" + index.toString());
        parentNode.removeChild(childNode);
    }

function getValue(name) {
        if (data.hasOwnProperty(name)) {
            data[name] = data[name] + 1
            let value = data[name]
            return name + "_div_" + value.toString()
        } else {
            let value = 1
            data[name] = 1
            return name + "_div_" + value.toString()
        }
    }

function addInput(name) {
        let div_class = getValue(name)
        let index = data[name].toString()
        let index_ = data[name]
        let main_div = document.getElementById(name)
        let div = document.createElement("div")
        div.className = name + "_div"
        div.id = div_class
        let input_sheet_name = document.createElement("input")
        input_sheet_name.type = "hidden"
        input_sheet_name.value = name
        input_sheet_name.name = name + "_sheetName_" + index
        input_sheet_name.className = "input"
        div.appendChild(input_sheet_name)
        let label_start_merge = document.createElement("label")
        label_start_merge.htmlFor = name + "_startMerge_" + index
        label_start_merge.innerHTML = "合并开始的行列:"
        div.appendChild(label_start_merge)
        let input_start_merge = document.createElement("input")
        input_start_merge.type = 'text'
        input_start_merge.name = name + "_startMerge_" + index
        input_start_merge.id = name + "_startMerge_" + index
        input_start_merge.className = "input"
        div.appendChild(input_start_merge)
        let label_end_merge = document.createElement("label")
        label_end_merge.htmlFor = name + "_endMerge_" + index
        label_end_merge.innerHTML = "合并结束的行列:"
        div.appendChild(label_end_merge)
        let input_end_merge = document.createElement("input")
        input_end_merge.type = 'text'
        input_end_merge.name = name + "_endMerge_" + index
        input_end_merge.id = name + "_endMerge_" + index
        input_end_merge.className = "input"
        div.appendChild(input_end_merge)
        let label_select_value = document.createElement("label")
        label_select_value.htmlFor = name + "_selectValue_" + index
        label_select_value.innerHTML = "此位置的值:"
        div.appendChild(label_select_value)
        let select_value = document.createElement("select")
        select_value.name = name + "_selectValue_" + index
        select_value.id = name + "_selectValue_" + index
        select_value.style.display = "inline-block"
        select_value.addEventListener('change', function () {
            handleSelectChange(event, name, index_)
        })
        let select_value_sign = document.createElement("option")
        select_value_sign.value = "sign"
        select_value_sign.innerHTML = "分拣标志"
        select_value.appendChild(select_value_sign)
        let select_value_num = document.createElement("option")
        select_value_num.value = "num"
        select_value_num.innerHTML = "被分拣的数量"
        select_value.appendChild(select_value_num)
        let select_value_sum = document.createElement("option")
        select_value_sum.value = "sum"
        select_value_sum.innerHTML = "某一列的和"
        select_value.appendChild(select_value_sum)
        let select_value_mean = document.createElement("option")
        select_value_mean.value = "mean"
        select_value_mean.innerHTML = "某一列的平均值"
        select_value.appendChild(select_value_mean)
        let select_value_median = document.createElement("option")
        select_value_median.value = "median"
        select_value_median.innerHTML = "某一列的中位数"
        select_value.appendChild(select_value_median)
        let select_value_min = document.createElement("option")
        select_value_min.value = "min"
        select_value_min.innerHTML = "某一列的最小值"
        select_value.appendChild(select_value_min)
        let select_value_max = document.createElement("option")
        select_value_max.value = "max"
        select_value_max.innerHTML = "某一列的最大值"
        select_value.appendChild(select_value_max)
        let select_value_user_defined = document.createElement("option")
        select_value_user_defined.value = "userDefined"
        select_value_user_defined.innerHTML = "自定义"
        select_value.appendChild(select_value_user_defined)
        div.appendChild(select_value)
        let label_value = document.createElement("label")
        label_value.htmlFor = name + "_value_" + index
        label_value.id = name + "_valueLabel_" + index
        label_value.innerHTML = "所选择的列:"
        label_value.style.display = "none"
        div.appendChild(label_value)
        let input_value = document.createElement("input")
        input_value.type = 'number'
        input_value.name = name + "_value_" + index
        input_value.id = name + "_value_" + index
        input_value.className = "input"
        input_value.style.display = "none"
        div.appendChild(input_value)
        let label_alignment = document.createElement("label")
        label_alignment.htmlFor = name + "_alignment_" + index
        label_alignment.innerHTML = "对齐方式:"
        div.appendChild(label_alignment)
        let select_alignment = document.createElement("select")
        select_alignment.name = name + "_alignment_" + index
        select_alignment.id = name + "_alignment_" + index
        let select_alignment_median = document.createElement("option")
        select_alignment_median.value = "median"
        select_alignment_median.innerHTML = "居中对齐"
        select_alignment.appendChild(select_alignment_median)
        let select_alignment_left = document.createElement("option")
        select_alignment_left.value = "left"
        select_alignment_left.innerHTML = "左对齐"
        select_alignment.appendChild(select_alignment_left)
        let select_alignment_right = document.createElement("option")
        select_alignment_right.value = "right"
        select_alignment_right.innerHTML = "右对齐"
        select_alignment.appendChild(select_alignment_right)
        div.appendChild(select_alignment)
        let label_font_size = document.createElement("label")
        label_font_size.htmlFor = name + "_fontSize_" + index
        label_font_size.innerHTML = "字体大小:"
        div.appendChild(label_font_size)
        let input_font_size = document.createElement("input")
        input_font_size.type = 'text'
        input_font_size.name = name + "_fontSize_" + index
        input_font_size.id = name + "_fontSize_" + index
        input_font_size.className = "input1"
        input_font_size.value = "24"
        div.appendChild(input_font_size)
        let label_font_type = document.createElement("label")
        label_font_type.htmlFor = name + "_fontType_" + index
        label_font_type.innerHTML = "字体:"
        div.appendChild(label_font_type)
        let input_font_type = document.createElement("input")
        input_font_type.type = 'text'
        input_font_type.name = name + "_fontType_" + index
        input_font_type.id = name + "_fontType_" + index
        input_font_type.className = "input2"
        input_font_type.value = "微软雅黑"
        div.appendChild(input_font_type)
        let input_delete = document.createElement("input")
        input_delete.type = "button"
        input_delete.value = "删除"
        input_delete.addEventListener('click', function () {
            deleteNode(name, index_)
        })
        div.appendChild(input_delete)
        main_div.appendChild(div)
    }