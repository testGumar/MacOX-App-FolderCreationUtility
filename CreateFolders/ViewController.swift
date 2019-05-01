//
//  ViewController.swift
//  CreateFolders
//
//  Created by Umar Sayyed on 30/04/19.
//  Copyright © 2019 Umar Sayyed. All rights reserved.
//

import Cocoa
import CoreXLSX
//import BRAOfficeDocumentPackage

class ViewController: NSViewController {
    
    fileprivate var excelPath: String = ""
    fileprivate var outputDirPath: String = ""
    fileprivate var jobNumber: String = ""
    fileprivate var schoolName: String = ""
    
    @IBOutlet weak var btnBrowse: NSButton!
    @IBOutlet weak var btnDone: NSButton!
    
    @IBOutlet weak var tfJobNumber: NSTextField!
    @IBOutlet weak var tfSchoolName: NSTextField!
    
    @IBOutlet weak var tfExcelPath: NSTextField!
    @IBOutlet weak var tfOutputDirPath: NSTextField!
    
    
    let PARENT_FOLDER_NAME = ""
    
    override func viewDidLoad() {
        super.viewDidLoad()

        // Do any additional setup after loading the view.
        
        
        
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
//        if (jobNumber.isEmpty) {
//            showAlert(title: "Job Number is required", msg: "Enter valid job number.")
//        } else if (schoolName.isEmpty) {
//            showAlert(title: "School Name is required", msg: "Enter valid school name.")
//        } else if (excelPath.isEmpty) {
//            showAlert(title: "Excel is required", msg: "Select valid excel.")
//        } else if (outputDirPath.isEmpty) {
//            showAlert(title: "Output Directory is required", msg: "Select valid output directory.")
//        } else {
//            processData()
//        }
        processData1()
    }
    
    func processData1 () {
        var spreadsheet: BRAOfficeDocumentPackage = BRAOfficeDocumentPackage.open(excelPath)
        if let array = spreadsheet.workbook.sheets as? Array<BRASheet> {
            for sheet in array {
                print(sheet.name)
            }
        }
        
        
    }
    
    func processData () {
        jobNumber = jobNumber.replacingOccurrences(of: " ", with: "_")
        
        guard let file = XLSXFile(filepath: excelPath) else {
            showAlert(title: "Invalid Excel File", msg: "Unable to open the selected excel file, try different one.")
            return;
        }
        
        do {
            for path in try file.parseWorksheetPaths() {
                let ws = try file.parseWorksheet(at: path)
                for row in ws.data?.rows ?? [] {
                    for c in row.cells {
                        print(c)
                    }
                }
            }
        } catch {
            print("Error: ", error)
        }
        
    }
    
    
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
