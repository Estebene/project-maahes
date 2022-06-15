use docx_rs::*;

use std::collections::HashMap;
use std::collections::hash_map::Entry;
use std::fs::File;
use std::io::{Read, Write};
use std::vec;


pub fn run() {
    let keywords = ["scc", "cemetery", "chp"];
    let mut paragraphs:HashMap<String, Vec<Paragraph>> = HashMap::new();
    let mut paragraph_buffer: Vec<Paragraph> = vec![];
    let mut file = File::open("./test.docx").unwrap();
    let mut buf = vec![];
    file.read_to_end(&mut buf).unwrap();
    
    let mut file = File::create("./hello.json").unwrap();
    let res = read_docx(&buf).unwrap();
    let res_json = res.json();
    // println!("{:?}", res);
    // for style in res.styles.iter() {
    //     println!("{:?}", res.styles);
    // } 
    let document_style = res.styles;


    for child in res.document.children.iter().rev() {
        let paragraph = match child {
            DocumentChild::Paragraph(value) => &**value,
            _ => panic!("not normal")
        };

        let paragraph_type = &paragraph.property.style.as_ref().unwrap().val; 

        if paragraph_type == "Normal" {
            paragraph_buffer.push(paragraph.clone());
            continue
        }

        if paragraph_type != "Heading1" {continue} 

        paragraph_buffer.push(paragraph.clone());

        let texts = match &paragraph.children[0] {
            ParagraphChild::Run(value) => &value.children, 
            _ => panic!("not normal")
        };

        if texts.len() > 0 {
            let text = match &texts[0] {
                RunChild::Text(value) => value,
                _ => panic!("not normal")
            };

            let lower_text = text.text.to_lowercase();
            let text_arr: Vec<&str> = lower_text.split(" ").collect();


            for keyword in keywords.iter() {
                if text_arr.contains(keyword){
                    for paragraph in (&paragraph_buffer).iter() {
                        match paragraphs.entry(keyword.to_string()) {
                            Entry::Vacant(e) => { e.insert(vec![paragraph.clone()]); },
                            Entry::Occupied(mut e) => { e.get_mut().push(paragraph.clone()); }
                        }
                    }
                    paragraph_buffer.clear();
                }
            }
            // println!("{:?}", text_arr)
        }
    }

    for (keyword, paragraph_list) in paragraphs {
        let path_str = format!("./{}.docx", keyword);
        let path = std::path::Path::new(&path_str);
        let file = std::fs::File::create(&path).unwrap();
        let mut doc = Docx::new();
        for paragraph in paragraph_list.iter().rev() {
            let paragraph_type = &paragraph.property.style.as_ref().unwrap().val;
            println!("{}",paragraph_type);
            doc = doc.add_paragraph(paragraph.clone().style(&paragraph_type.as_str()));
        }

        doc.styles(document_style.clone()).build().pack(file).ok();
    }


    file.write_all(res_json.as_bytes()).unwrap();
    file.flush().unwrap();
}