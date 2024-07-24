// use std::borrow::Cow;

use std::borrow::Cow;

use docx_rust::document;
use fancy_regex::{Captures, Regex};

fn main() {
    use docx_rust::DocxFile;
    use hard_xml::{XmlRead, XmlWrite};
    use tera::Tera;

    let docx = DocxFile::from_file("template_with_table.docx").unwrap();
    let mut docx = docx.parse().unwrap();
    let document: document::Document = docx.document;
    // let body = docx.document.body;

    let mut src_xml = XmlWrite::to_string(&document).unwrap();
    let re = Regex::new(r"(?<={)(<[^>]*>)+(?=[\{%\#])|(?<=[%\}\#])(<[^>]*>)+(?=\})").unwrap();
    src_xml = re.replace_all(&src_xml, "").parse().unwrap();

    let re = Regex::new(r"{%(?:(?!%}).)*|{#(?:(?!#}).)*|{{(?:(?!}}).)*").unwrap();
    src_xml = re
        .replace_all(&src_xml, |caps: &Captures| {
            let re = Regex::new(r"</w:t>.*?(<w:t>|<w:t [^>]*>)").unwrap();
            re.replace_all(&caps[0], "").into_owned()
        })
        .parse()
        .unwrap();
   
    // manage table cell colspan
    let re = Regex::new(r"(<w:tc[ >](?:(?!<w:tc[ >]).)*){%\s*colspan\s+([^%]*)\s*%}(.*?</w:tc>)")
        .unwrap();
    src_xml = re
        .replace_all(&src_xml, |caps: &Captures| {
            let mut cell_xml = format!("{}{}", &caps[1], &caps[3]);
            let re = Regex::new(r"<w:r[ >](?:(?!<w:r[ >]).)*<w:t></w:t>.*?</w:r>").unwrap();
            cell_xml = re.replace_all(&cell_xml, "").parse().unwrap();

            let re = Regex::new(r"<w:gridSpan[^/]*/>").unwrap();
            cell_xml = re.replace(&cell_xml, "").parse().unwrap();

            let re = Regex::new(r"(<w:tcPr[^>]*>)").unwrap();
            re.replace_all(&cell_xml, |child_caps:&Captures|{
                let rep = format!(
                    r#"{}<w:gridSpan w:val="{{{}}}"/>"/>"#, &child_caps[1],
                    &caps[2]
                );
                Cow::Owned(String::from(rep))
            }).into_owned()
        })
        .parse()
        .unwrap();

    //manage table cell background color
    let re = Regex::new(r"(<w:tc[ >](?:(?!<w:tc[ >]).)*){%\s*cellbg\s+([^%]*)\s*%}(.*?</w:tc>)")
        .unwrap();
    src_xml = re
        .replace_all(&src_xml, |caps: &Captures| {
            let mut cell_xml = format!("{}{}", &caps[1], &caps[3]);
            let re = Regex::new(r"<w:r[ >](?:(?!<w:r[ >]).)*<w:t></w:t>.*?</w:r>").unwrap();
            cell_xml = re.replace_all(&cell_xml, "").parse().unwrap();

            let re = Regex::new(r"<w:shd[^/]*/>").unwrap();
            cell_xml = re.replace(&cell_xml, "").parse().unwrap();

            let re = Regex::new(r"(<w:tcPr[^>]*>)").unwrap();
            re.replace_all(&cell_xml, |child_caps:&Captures|{
                let rep = format!(
                    r#"{}<w:shd w:val="clear" w:color="auto" w:fill="{{{}}}"/>"#, &child_caps[1],
                    &caps[2]
                );
                Cow::Owned(String::from(rep))
            }).into_owned()
        })
        .parse()
        .unwrap();

    // ensure space preservation
    let re = Regex::new(r"<w:t>((?:(?!<w:t>).)*)({{.*?}}|{%.*?%})").unwrap();
    src_xml = re.replace_all(&src_xml, |caps: &Captures| {
        let rep = format!(r#"<w:t xml:space="preserve">{}{}"#, &caps[1],&caps[2]);
        Cow::Owned(String::from(rep))
    })
    .parse()
    .unwrap();

    let re = Regex::new(r"({{r\s.*?}}|{%r\s.*?%})").unwrap();
    src_xml = re.replace_all(&src_xml, |caps: &Captures| {
        let rep = format!(r#"</w:t></w:r><w:r><w:t xml:space="preserve">{}</w:t></w:r><w:r><w:t xml:space="preserve">"#, &caps[1]);
        Cow::Owned(String::from(rep))
    })
    .parse()
    .unwrap();

    //{%- will merge with previous paragraph text
    let re = Regex::new(r"</w:t>(?:(?!</w:t>).)*?{%-").unwrap();
    src_xml = re.replace_all(&src_xml, "{%").parse().unwrap();
    //-%} will merge with next paragraph text
    let re = Regex::new(r"-%}(?:(?!<w:t[ >]|{%|{{).)*?<w:t[^>]*?>").unwrap();
    src_xml = re.replace_all(&src_xml, "%}").parse().unwrap();

    let v = vec!["tr", "tc", "p", "r"];
    for it in &v {
        let pat = format!(
            "<w:{y}[ >](?:(?!<w:{y}[ >]).)*({{%|{{{{){y} ([^}}%]*(?:%}}|}}}})).*?</w:{y}>",
            y = it
        );
        let re = Regex::new(&pat).unwrap();
        src_xml = re.replace_all(&src_xml, |caps: &Captures| {
            let rep = format!("{}{}", &caps[1],&caps[2]);
            Cow::Owned(String::from(rep))
        })
        .parse()
        .unwrap();
    }

    let v = vec!["tr", "tc", "p"];
    for it in &v {
        let pat = format!(
            "<w:{y}[ >](?:(?!<w:{y}[ >]).)*({{#){y} ([^}}#]*(?:#}})).*?</w:{y}>",
            y = it
        );
        let re = Regex::new(&pat).unwrap();
        src_xml = re.replace_all(&src_xml, |caps: &Captures| {
            let rep = format!("{}{}", &caps[1],&caps[2]);
            Cow::Owned(String::from(rep))
        })
        .parse()
        .unwrap();
    }

    let re = Regex::new(r"(?<=\{[\{%])(.*?)(?=[\}%]})").unwrap();
    src_xml = re
        .replace_all(&src_xml, |caps: &Captures| {
            let tags = &caps[0]
                .replace(r"&#8216;", "")
                .replace("&lt;", "<")
                .replace("&gt;", ">")
                .replace("“", "\"")
                .replace("”", "\"")
                .replace("‘", "'")
                .replace("’", "'");
            Cow::Owned(String::from(tags))
        })
        .parse()
        .unwrap();

    let mut tera = Tera::default();
    tera.add_raw_template("input_text", &src_xml).unwrap();
    // Prepare the context with some data
    let mut context = tera::Context::new();
    context.insert("bridgeName", "World");
    context.insert("projectName", "测试项目来打发");

    let mut data_list = Vec::new();
    let v = vec!["数据1A", "数据1B", "数据1C"];
    data_list.push(v);
    let t = format!("{{% selfvm_first %}}{}","数据2B");
    let v = vec!["数据2A", &t, "数据2C"];
    data_list.push(v);
    let t = format!("{{% selfvm %}}{}","数据2B");
    let v = vec!["数据3A",  &t, "数据3C"];
    data_list.push(v);
    let v = vec!["数据4A",  &t, "数据4C"];
    data_list.push(v);
    let v = vec!["数据5A", "数据5B", "数据5C"];
    data_list.push(v);
    context.insert("dataList", &data_list);

    // 填充文本数据
    src_xml = tera.render("input_text", &context).unwrap();

    // 合并单元格
    let re = Regex::new(r"<w:tc[ >](?:(?!<w:tc[ >]).)*?{%\s*selfvm_first\s*%}.*?</w:tc[ >]").unwrap();
    src_xml = re
        .replace_all(&src_xml, |caps: &Captures| {
            let re = Regex::new(r"(</w:tcPr[ >].*?<w:t(?:.*?)>)(.*?)(?:{%\s*selfvm_first\s*%})(.*?)(</w:t>)").unwrap();
           let c = re.replace_all(&caps[0], |child_caps: &Captures| {
                let mut rep = String::from(r#"<w:vMerge w:val="restart"/>"#);
                rep.push_str(&child_caps[1]);
                rep.push_str(&child_caps[2]);
                rep.push_str(&child_caps[3]);
                rep.push_str(&child_caps[4]);
                Cow::Owned(rep)
            });
            dbg!(&c);
            c.into_owned()
        })
        .parse()
        .unwrap();    
    let re = Regex::new(r"<w:tc[ >](?:(?!<w:tc[ >]).)*?{%\s*selfvm\s*%}.*?</w:tc[ >]").unwrap();
    src_xml = re
        .replace_all(&src_xml, |caps: &Captures| {
            let re = Regex::new(r"(</w:tcPr[ >].*?<w:t(?:.*?)>)(.*?)(?:{%\s*selfvm\s*%})(.*?)(</w:t>)").unwrap();
            re.replace_all(&caps[0], |child_caps: &Captures| {
                let mut rep = String::from(r#"<w:vMerge w:val="continue"/>"#);
                rep.push_str(&child_caps[1]);
                rep.push_str(&child_caps[4]);
                Cow::Owned(rep)
            }).into_owned()
        })
        .parse()
        .unwrap();    
    println!("{}",&src_xml);
    // docx_rust 丢失合并代码<w:vMerge /> 需要修改docx_rust中的docx_rust::formatting::TableCellProperty
    docx.document = XmlRead::from_str(&src_xml).unwrap();
   
    docx.write_file("out.docx").unwrap();
}
