//
//  UserData.swift
//  Feez
//
//  Created by Mehul Parmar on 28/02/20.
//  Copyright © 2020 CROISSANCE 3W. All rights reserved.
//

import Foundation


class GenerateTrip {
    static let shared = GenerateTrip()
    
    func getTrips(_ tripCount: Int) -> [TripDetail]{
        let trip1 = TripDetail.init(serialNumber: 1, date: Date().adjust(hour: 8, minute: 0, second: 0, day: 1, month: 2), kms: 0.10, miles: 0.06, departure: AddressTrip.init(city: "Ahmedabad", address: "Gota Ahmedabad"), arrival: AddressTrip.init(city: "Ahmedabad", address: "Gota Ahmedabad"))
        let trip2 = TripDetail.init(serialNumber: 2, date: Date().adjust(hour: 9, minute: 0, second: 0, day: 2, month: 2), kms: 0.15, miles: 0.11, departure: AddressTrip.init(city: "Ahmedabad", address: "Gota Ahmedabad"), arrival: AddressTrip.init(city: "Ahmedabad", address: "Gota Ahmedabad"))
        let trip3 = TripDetail.init(serialNumber: 3, date: Date().adjust(hour: 10, minute: 0, second: 0, day: 3, month: 2), kms: 0.5, miles: 0.1, departure: AddressTrip.init(city: "Ahmedabad", address: "Gota Ahmedabad"), arrival: AddressTrip.init(city: "Ahmedabad", address: "Gota Ahmedabad"))
        let trip4 = TripDetail.init(serialNumber: 4, date: Date().adjust(hour: 11, minute: 0, second: 0, day: 4, month: 2), kms: 0.20, miles: 0.16, departure: AddressTrip.init(city: "Ahmedabad", address: "Gota Ahmedabad"), arrival: AddressTrip.init(city: "Ahmedabad", address: "Gota Ahmedabad"))
        var arrTrip = [TripDetail]()
        arrTrip.append(trip1)
        arrTrip.append(trip2)
        arrTrip.append(trip3)
        arrTrip.append(trip4)
        
        return arrTrip
    }
}

struct TripDetail : Codable {
    var serialNumber : Int
    var date: Date
    var kms: Double
    var miles: Double

    var departure   : AddressTrip
    var arrival     : AddressTrip

    enum CodingKeys: String, CodingKey {
        case serialNumber = "serialNumber"
        case date = "date"
        
        case departure = "departure"
        case arrival = "arrival"
        
        case kms = "kms"
        case miles = "miles"
    }
}

struct AddressTrip : Codable {
    var city : String
    var address : String
    
    enum CodingKeys: String, CodingKey {
        case city = "city"
        case address = "address"
    }
}

