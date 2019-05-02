//
//  ViewController.swift
//  CreateFolders
//
//  Created by Umar Sayyed on 30/04/19.
//  Copyright © 2019 Umar Sayyed. All rights reserved.
//

import Cocoa
//import CoreXLSX

class ViewController: NSViewController {
    
    fileprivate var excelPath: String = ""
    fileprivate var outputDirPath: String = ""
    fileprivate var classFolderPath: String = ""
    fileprivate var studentFolderPath: String = ""
    fileprivate var jobNumber: String = ""
    fileprivate var schoolName: String = ""
    
    @IBOutlet weak var btnBrowse: NSButton!
    @IBOutlet weak var btnDone: NSButton!
    
    @IBOutlet weak var tfJobNumber: NSTextField!
    @IBOutlet weak var tfSchoolName: NSTextField!
    
    @IBOutlet weak var tfExcelPath: NSTextField!
    @IBOutlet weak var tfOutputDirPath: NSTextField!
    @IBOutlet weak var tfError: NSTextField!
    @IBOutlet weak var progress: NSProgressIndicator!
    
    
    let PARENT_FOLDER_NAME = ""
    let filemgr = FileManager.default
    
    let FOLDER_NAMES = ["Class", "Staff", "Marketing", "Originals", "Students", "Raws", "Siblings"]
    let SUB_FOLDER_NAMES = ["Selects", "QRCodes", "StudentNames"]
    
    override func viewDidLoad() {
        super.viewDidLoad()

        // Do any additional setup after loading the view.
        progress.style = .spinning
        
        
    }
    @IBAction func outputAction(_ sender: Any) {   let panel = NSOpenPanel()
        panel.canChooseFiles = false;
        panel.canChooseDirectories = true;
        panel.allowsMultipleSelection = false;
        
        if panel.runModal() == .OK {
            for file in panel.urls {
                outputDirPath = file.path
                tfOutputDirPath.stringValue = outputDirPath
                print("file path - ", file.path)
            }
        }
        
    }
    
    @IBAction func browseAction(_ sender: Any) {
        
        let panel = NSOpenPanel()
        panel.canChooseFiles = true;
        panel.canChooseDirectories = false;
        panel.allowedFileTypes = ["xl","xls", "xlsx"]
        panel.allowsMultipleSelection = false;
        
        if panel.runModal() == .OK {
            for file in panel.urls {
                excelPath = file.path
                tfExcelPath.stringValue = excelPath
                print("file path - ", file.path)
            }
        }
    }
    
    @IBAction func doneAction(_ sender: Any) {
        jobNumber = tfJobNumber.stringValue.trim().uppercased()
        schoolName = tfSchoolName.stringValue.trim()
        if (jobNumber.isEmpty) {
            showAlert(title: "Job Number is required", msg: "Enter valid job number.")
        } else if (schoolName.isEmpty) {
            showAlert(title: "School Name is required", msg: "Enter valid school name.")
        } else if (excelPath.isEmpty) {
            showAlert(title: "Excel is required", msg: "Select valid excel file.")
        } else if (outputDirPath.isEmpty) {
            showAlert(title: "Output Directory is required", msg: "Select output directory.")
        } else {
            
            progress.isHidden = false
            
            jobNumber = jobNumber.trim().replaceSpacewithUnderscore()
            schoolName = schoolName.trim().replaceSpacewithUnderscore()
            outputDirPath = outputDirPath+"/"+jobNumber+"_"+schoolName;
            processData1()
            
            DispatchQueue.main.async {
                self.progress.isHidden = true;
            }
        }
    }
    
    func processData1 () {
        //KADEV2019007
        //EuroKids
        
        createOuterFolders()
        
        let spreadsheet: BRAOfficeDocumentPackage = BRAOfficeDocumentPackage.open(excelPath)
        guard let worksheets = spreadsheet.workbook.worksheets as? Array<BRAWorksheet> else { return }
        for ws in worksheets {
            guard let rows = ws.rows as? Array<BRARow> else { return }
            guard let cols = ws.columns as? Array<BRAColumn> else { return }
            print("rows count = ", rows.count)
            print("cols count = ", cols.count)
            if rows.count <= 2 && cols.count <= 1 {
                return
            }
            
            for r in 2...rows.count {
                var path = ""
                for c in 1...3 {
                    guard let col_index = BRAColumn.columnName(forColumnIndex: c) else {continue}
                    guard var content = ws.cell(forCellReference: "\(col_index)\(r)")?.stringValue() else {continue}
//                    print(content)
                    if !content.isEmpty && content.count >= 3 {
                        content = content.trim().replaceSpacewithUnderscore()
                        path.append("/"+jobNumber+"_"+content)
                    }
                    
                    if (c == 1) {
                        do {
                            print("path", outputDirPath + path)
                            try filemgr.createDirectory(atPath: classFolderPath + path, withIntermediateDirectories: true, attributes: nil)
                        } catch let er {
                            print(er)
                            tfError.stringValue = tfError.stringValue + "\n" + er.localizedDescription
                        }
                    }
                }
                
                do {
                    print("path", outputDirPath + path)
                    try filemgr.createDirectory(atPath: studentFolderPath + path, withIntermediateDirectories: true, attributes: nil)
                } catch let er {
                    print(er)
                    tfError.stringValue = tfError.stringValue + "\n" + er.localizedDescription
                }
                //break
            }
            //break
        }
        
       
        
    }
    
    func createOuterFolders () {
        
        
        for folder in FOLDER_NAMES {
            let newPath = outputDirPath + "/" + jobNumber + "_" + folder
            
            if folder == "Students" {
                for subFolder in SUB_FOLDER_NAMES {
                    let subPath =  newPath + "/" + jobNumber + "_" + subFolder
                    do {
                        try filemgr.createDirectory(atPath: subPath, withIntermediateDirectories: true, attributes: nil)
                    } catch let er {
                        print(er)
                        tfError.stringValue = tfError.stringValue + "\n" + er.localizedDescription
                    }
                    if subFolder == "StudentNames" {
                        studentFolderPath = subPath
                    }
                }
                continue
            }
            do {
                try filemgr.createDirectory(atPath: newPath, withIntermediateDirectories: true, attributes: nil)
            } catch let er {
                print(er)
                tfError.stringValue = tfError.stringValue + "\n" + er.localizedDescription
            }
            
            if folder == "Class" {
                classFolderPath = newPath
            }
        }
        
//        do {
//            print("path", outputDirPath + path)
//            try filemgr.createDirectory(atPath: outputDirPath + path, withIntermediateDirectories: true, attributes: nil)
//        } catch let er { print(er) }
    }
    
//    func processData () {
//        jobNumber = jobNumber.trim().replaceSpacewithUnderscore()
//        schoolName = schoolName.trim().replaceSpacewithUnderscore()
//
//        guard let file = XLSXFile(filepath: excelPath) else {
//            showAlert(title: "Invalid Excel File", msg: "Unable to open the selected excel file, try different one.")
//            return;
//        }
//
//        do {
//            for path in try file.parseWorksheetPaths() {
//                let ws = try file.parseWorksheet(at: path)
//                for row in ws.data?.rows ?? [] {
//                    for c in row.cells {
//                        print(c)
//                    }
//                }
//            }
//        } catch {
//            print("Error: ", error)
//        }
//
//    }
    
    
    override var representedObject: Any? {
        didSet {
        // Update the view, if already loaded.
        }
    }
    
    func showAlert (title:String, msg:String) {
        let alert = NSAlert.init()
        alert.messageText = title
        alert.informativeText = msg
        alert.addButton(withTitle: "OK")
        alert.addButton(withTitle: "Cancel")
        alert.runModal()
    }


}


extension String {
    func replaceSpacewithUnderscore () -> String {
        return self.replacingOccurrences(of: " ", with: "_")
    }
    func trim () -> String {
        return self.trimmingCharacters(in: .whitespacesAndNewlines)
    }
}


extension Data {
    func toDict () -> Dictionary<String, Any>? {
        return NSKeyedUnarchiver.unarchiveObject(with: self as Data) as? Dictionary<String, Any>
    }
    
    func toArray () -> Array<Any>? {
        return NSKeyedUnarchiver.unarchiveObject(with: self as Data) as? Array<Any>
    }
}

extension Array {
    func toJSON() -> String {
        guard let data = try? JSONSerialization.data(withJSONObject: self, options: []) else {
            return ""
        }
        return String(data: data, encoding: String.Encoding.utf8)!
    }
}

extension Dictionary {
    
    func toJSON () -> String {
        let invalidJson = "Not a valid JSON"
        do {
            let jsonData = try JSONSerialization.data(withJSONObject: self, options: .prettyPrinted)
            return String(bytes: jsonData, encoding: String.Encoding.utf8) ?? invalidJson
        } catch {
            return invalidJson
        }
    }
    
    var json: String {
        let invalidJson = "Not a valid JSON"
        do {
            let jsonData = try JSONSerialization.data(withJSONObject: self, options: .prettyPrinted)
            return String(bytes: jsonData, encoding: String.Encoding.utf8) ?? invalidJson
        } catch {
            return invalidJson
        }
    }
}
