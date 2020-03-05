//
//  ViewController.swift
//  Test
//
//  Created by Mehul Parmar on 03/03/20.
//  Copyright 2013-2020, John McNamara.. All rights reserved.
//

import UIKit
import xlsxwriter
import QuickLook
import Foundation

class ViewController: UIViewController {
    
    //URL Path for physical location of excell sheet file
    var url: URL? = nil
    let preview = QLPreviewController()
    
    override func viewDidLoad() {
        super.viewDidLoad()
        
        //Generate custom sheet
        createSheet()
        
        //self.btnViewSheetClicked(UIButton())
    }
    
    // MARK: Create Sheet
    func createSheet() {
        //get path to save file
        let path = getPath(createFile: "Daily Export")
        
        //Generate workbook
        let workbook = workbook_new(path)
        
        //Add one Sheet in workbook //You can add multiple sheet
        let worksheet = workbook_add_worksheet(workbook, "Daily Export");
        
        //Format : Means set properties to row, column & cell
        let formatHeader = setupFormatHeader(using: workbook)
        setupHeader(using: worksheet, myformatBold: formatHeader)
        let formatRightAllign_Bold = setupFormatRightAllign_Bold(using: workbook)
        let formatAllignCenter = setupFormatCenterAllign(using: workbook)
        let formatNumber = setupNumberFormat(using: workbook)
        let formatText = setupFormatText(using: workbook)
        
        //Get dymmy data to display it in Sheet
        let arrTrips = GenerateTrip.shared.getTrips(10)
        for (index, trip) in arrTrips.enumerated() {
            
            //Set Serial number
            let serialNumber = index + 2
            let row = lxw_row_t(serialNumber)
            
            //Set custom width to column
            worksheet_set_row(worksheet, 0, 30, formatAllignCenter)
            worksheet_set_column(worksheet, 0, 0, 8, formatAllignCenter)
            worksheet_set_column(worksheet, 1, 1, 30, nil)
            worksheet_set_column(worksheet, 2, 2, 20, nil)
            worksheet_set_column(worksheet, 3, 3, 30, nil)
            
            worksheet_set_column(worksheet, 4, 4, 20, nil)
            worksheet_set_column(worksheet, 5, 5, 30, nil)
            worksheet_set_column(worksheet, 6, 6, 8, formatAllignCenter)
            worksheet_set_column(worksheet, 7, 7, 8, formatAllignCenter)
            
            //Set Data/Value in Cell
            //For number, we use : worksheet_write_number
            //For String, we use : worksheet_write_string
            worksheet_write_number(worksheet, row, 0, Double(trip.serialNumber), nil)
            worksheet_write_string(worksheet, row, 1, "\(trip.date.toString(style: DateStyleType.long))", formatText)
            
            //Deprture
            worksheet_write_string(worksheet, row, 2, "\(trip.departure.city)", formatText)
            worksheet_write_string(worksheet, row, 3, "\(trip.departure.address)", formatText)
            
            //Arrival
            worksheet_write_string(worksheet, row, 4, "\(trip.arrival.city)", formatText)
            worksheet_write_string(worksheet, row, 5, "\(trip.arrival.address)", formatText)
            
            worksheet_write_number(worksheet, row, 6, trip.kms, formatNumber)
            worksheet_write_number(worksheet, row, 7, trip.miles, formatNumber)
            
            //Set sum
            let sumOfKMS = arrTrips.map({ $0.kms}).reduce(0.0, +)
            let sumOfMiles = arrTrips.map({ $0.miles}).reduce(0.0, +)
            worksheet_write_number(worksheet, row+1, 6, sumOfKMS, formatNumber)
            worksheet_write_number(worksheet, row+1, 7, sumOfMiles, formatNumber)
        }
        
        //Merge cell
        let rowTotal = lxw_row_t(arrTrips.count+2)
        let colTotal = lxw_col_t(5)
        worksheet_merge_range(worksheet, rowTotal, lxw_col_t(0), rowTotal, colTotal, "Total", formatRightAllign_Bold)
        
        //Save and close editing & generate physical file in document directory, If already exist then It will replace it
        workbook_close(workbook)
    }
    
    func getPath(createFile fileName: String) -> UnsafePointer<Int8> {
        url = FileManager.default.urls(for: .documentDirectory, in: .userDomainMask)[0].appendingPathComponent(fileName).appendingPathExtension("xlsx")
        let path = NSString(string: url!.path).fileSystemRepresentation
        print("path:", url!.absoluteString)
        return path
    }
    
    //IBActions for button 'View Sheet'
    @IBAction func btnViewSheetClicked(_ sender: UIButton) {
        DispatchQueue.main.asyncAfter(deadline: .now() + 0.3) {
            self.preview.dataSource = self
            self.preview.currentPreviewItemIndex = 0
            self.present(self.preview, animated: true)
        }
    }
}

// MARK: Set theme, Text, Allignment, Width, Style, Etc using lxw_format
extension ViewController {
    func setupFormatText(using workbook: UnsafeMutablePointer<lxw_workbook>?) -> UnsafeMutablePointer<lxw_format>? {
        let myformatNormal = workbook_add_format(workbook)
        format_set_align(myformatNormal, UInt8(LXW_ALIGN_VERTICAL_DISTRIBUTED.rawValue))
        format_set_text_wrap(myformatNormal);
        return myformatNormal
    }
    
    func setupFormatHeader(using workbook: UnsafeMutablePointer<lxw_workbook>?) -> UnsafeMutablePointer<lxw_format>? {
        let format = workbook_add_format(workbook)
        format_set_bold(format)
        format_set_align(format, UInt8(LXW_ALIGN_CENTER_ACROSS.rawValue))
        return format
    }
    
    func setupFormatRightAllign_Bold(using workbook: UnsafeMutablePointer<lxw_workbook>?) -> UnsafeMutablePointer<lxw_format>? {
        let format = workbook_add_format(workbook)
        format_set_bold(format)
        format_set_align(format, UInt8(LXW_ALIGN_RIGHT.rawValue))
        return format
    }
    
    func setupFormatCenterAllign(using workbook: UnsafeMutablePointer<lxw_workbook>?) -> UnsafeMutablePointer<lxw_format>? {
        let format = workbook_add_format(workbook)
        //        format_set_align(format, UInt8(LXW_ALIGN_CENTER.rawValue))
        //        format_set_align(format, UInt8(LXW_ALIGN_VERTICAL_CENTER.rawValue))
        format_set_align(format, UInt8(LXW_ALIGN_CENTER_ACROSS.rawValue))
        return format
    }
    
    func setupNumberFormat(using workbook: UnsafeMutablePointer<lxw_workbook>?) -> UnsafeMutablePointer<lxw_format>? {
        let formatDouble = workbook_add_format(workbook)
        format_set_num_format(formatDouble, "0.00");
        return formatDouble
    }
    
    func setupHeader(using
        worksheet: UnsafeMutablePointer<lxw_worksheet>?,
                     myformatBold: UnsafeMutablePointer<lxw_format>?) {
        //Set Columb
        worksheet_write_string(worksheet, 0, 0, "Serial No.", myformatBold)
        worksheet_write_string(worksheet, 0, 1, "Date", myformatBold)
        worksheet_write_string(worksheet, 0, 2, "Departure", myformatBold)
        worksheet_write_string(worksheet, 0, 3, "", myformatBold)
        worksheet_write_string(worksheet, 0, 4, "Arrival", myformatBold)
        worksheet_write_string(worksheet, 0, 5, "", myformatBold)
        worksheet_write_string(worksheet, 1, 6, "Kms", myformatBold)
        worksheet_write_string(worksheet, 1, 7, "Miles", myformatBold)
        
        worksheet_write_string(worksheet, 1, 2, "City", myformatBold)
        worksheet_write_string(worksheet, 1, 3, "Address", myformatBold)
        worksheet_write_string(worksheet, 1, 4, "City", myformatBold)
        worksheet_write_string(worksheet, 1, 5, "Address", myformatBold)
        
        worksheet_merge_range(worksheet, 0, 0, 1, 0, "Serial No.", myformatBold)
        worksheet_merge_range(worksheet, 0, 1, 1, 1, "Date", myformatBold)
        
        worksheet_merge_range(worksheet, 0, 6, 1, 6, "Kms", myformatBold)
        worksheet_merge_range(worksheet, 0, 7, 1, 7, "Miles", myformatBold)
        
        //worksheet_merge_range(worksheet, 0, 2, 0, 3, "Departure", myformatBold)
        //worksheet_merge_range(worksheet, 0, 4, 0, 5, "Arrival", myformatBold)
        
    }
}

// MARK: View Sheet
extension ViewController: QLPreviewControllerDataSource {
    func numberOfPreviewItems(in controller: QLPreviewController) -> Int {
        return 1
    }
    
    func previewController(_ controller: QLPreviewController, previewItemAt index: Int) -> QLPreviewItem {
        return url! as QLPreviewItem
    }
}


extension UIViewController {
    func exportFile(url: URL) {
        let activityVC = UIActivityViewController(activityItems: [url], applicationActivities: [])
        activityVC.excludedActivityTypes = [
            .message,
            .mail,
            .print,
            //             .copyToPasteboard
        ]
        activityVC.completionWithItemsHandler = {
            (activity, success, items, error) in
            if success {
                self.showSimpleAlert(message: "File export success")
            } else if let e = error {
                self.showSimpleAlert(message: "Export failed\n\(e)")
            } else {
                // Cancel export
            }
        }
        
        present(activityVC, animated: true)
    }
}

extension UIViewController {
    func showSimpleAlert(message: String?) {
        self.showSimpleAlert(title: "Notice", message: message)
    }
    
    func showSimpleAlert(title: String?, message: String?) {
        let alert = UIAlertController(title: title, message: message, preferredStyle: .alert)
        alert.addAction(UIAlertAction(title: "Confirm", style: .default, handler: nil))
        self.present(alert, animated: true, completion: nil)
    }
}
