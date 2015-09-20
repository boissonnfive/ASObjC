//
//  main.m
//  Facture
//
//  Created by Bruno Boissonnet on 07/09/2014.
//  Copyright (c) 2014 BInfoService. All rights reserved.
//

#import <Cocoa/Cocoa.h>

#import <AppleScriptObjC/AppleScriptObjC.h>

int main(int argc, const char * argv[])
{
    [[NSBundle mainBundle] loadAppleScriptObjectiveCScripts];
    return NSApplicationMain(argc, argv);
}
