﻿/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/**
 * Get value for key
 * @customfunction
 * @param key The key
 * @returns The value for the key.
 */
function getValueForKeyCF(key) {
  debugger;
  return key;
}

/**
 * Get value for key
 * @customfunction
 * @param key The key
 * @returns The value for the key.
 */
function setValueForKeyCF(key, value) {
  setValueForKey(key, value);
  return "Stored key/value pair";
}



/**
 * Add two numbers
 * @customfunction
 * @param key The key
 * @returns The value for the key.
 */
function add(a, b) {
  debugger;
  return a + b;
}


function dim(dimension, filters) {
  debugger
  return dimension
}



CustomFunctions.associate("GETVALUEFORKEYCF", getValueForKeyCF);
CustomFunctions.associate("SETVALUEFORKEYCF",setValueForKeyCF);
CustomFunctions.associate("ADD",add);
CustomFunctions.associate("DIM",dim);