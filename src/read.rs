use docx_rs::*;

use std::collections::HashMap;
use std::collections::hash_map::Entry;
use std::fs::File;
use std::io::{Read};
use std::path::PathBuf;
use std::vec;


pub fn run(path: PathBuf, folder_path: PathBuf, keywords: Vec<String>) {
    let mut paragraphs:HashMap<String, Vec<Paragraph>> = HashMap::new();
    let mut paragraph_buffer: Vec<Paragraph> = vec![];
    let mut file = File::open(path).unwrap();
    let mut buf = vec![];
    file.read_to_end(&mut buf).unwrap();
    
    // let mut file = File::create("./hello.json").unwrap();
    let res = read_docx(&buf).unwrap();
    // let res_json = res.json();
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

        if paragraph.property.style == None {
            paragraph_buffer.push(paragraph.clone());
            continue
        }

        let paragraph_type = &paragraph.property.style.as_ref().unwrap().val; 

        if paragraph_type.contains("Normal") {
            paragraph_buffer.push(paragraph.clone());
            continue
        }

        if !paragraph_type.contains("Heading") {continue} 


        if paragraph.children.len() == 0 {
            println!("empty paragraph");
            continue
        }

        // loop through all the runs in the paragraph creating a single string of text
        let mut texts: Vec<&RunChild> = vec![];
        for run in paragraph.children.iter() {
            match run {
                ParagraphChild::Run(value) => {
                    for text in value.children.iter() {
                        texts.push(text);
                    }
                },
                _ => panic!("not normal")
            }
        }
        let mut string: String = String::new();
        for text_child in texts.iter() {
            match text_child {
                RunChild::Text(value) => {
                    string.push_str(&value.text);
                },
                _ => panic!("not normal")
            }
        }

        println!("{:?}", string);


        if string.len() > 0 {
            paragraph_buffer.push(paragraph.clone());

            let lower_text = string.to_lowercase();
            let text_arr: Vec<&str> = lower_text.split(" ").collect();


            for keyword in keywords.iter() {
                if text_arr.contains(&keyword.as_str()){
                    for paragraph in (&paragraph_buffer).iter() {
                        match paragraphs.entry(keyword.to_string()) {
                            Entry::Vacant(e) => { e.insert(vec![paragraph.clone()]); },
                            Entry::Occupied(mut e) => { e.get_mut().push(paragraph.clone()); }
                        }
                    }
                }
            }
            paragraph_buffer.clear();
            // println!("{:?}", text_arr)
        }
    }

    for (keyword, paragraph_list) in paragraphs {
        let path_str = format!("{}/{}.docx", folder_path.to_string_lossy(), keyword);
        let path = std::path::Path::new(&path_str);
        let file = std::fs::File::create(&path).unwrap();
        let mut doc = Docx::new();
        for paragraph in paragraph_list.iter().rev() {
            if paragraph.property.style != None {
                let paragraph_type = &paragraph.property.style.as_ref().unwrap().val;
                doc = doc.add_paragraph(paragraph.clone().style(&paragraph_type.as_str()));
            }
            else {
                doc = doc.add_paragraph(paragraph.clone());
            }
        }

        doc.styles(document_style.clone()).build().pack(file).ok();
    }


}