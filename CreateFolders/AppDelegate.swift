//
//  AppDelegate.swift
//  CreateFolders
//
//  Created by Umar Sayyed on 30/04/19.
//  Copyright © 2019 Umar Sayyed. All rights reserved.
//

import Cocoa
import Fabric
import Crashlytics

@NSApplicationMain
class AppDelegate: NSObject, NSApplicationDelegate {



    func applicationDidFinishLaunching(_ aNotification: Notification) {
        // Insert code here to initialize your application
        UserDefaults.standard.register(defaults: ["NSApplicationCrashOnExceptions": true])
        Fabric.with([Crashlytics.self])
    }

    func applicationWillTerminate(_ aNotification: Notification) {
        // Insert code here to tear down your application
    }


}

