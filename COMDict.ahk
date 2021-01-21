Class COMDict {


    /*
        * Method: __New(keyvalArr := "")
        *   Creates a new instace of the dictionary class
        * Params:
        *   keyvalArr:  a dictionary creation formatted array
        *               defined below the definition of this class
        * Return:
        *   new dictionary instance
    */
    __New(keyvalArr := ""){
        this.dict := ComObjCreate("Scripting.Dictionary")
        this._setFromArr(keyvalArr)
    }


    /*
        * Method: setAnew(keyvalArr)
        *   sets the dict anew from the handed array
        * Params:
        *   keyvalArr:  a dictionary creation formatted array
        *               defined below the definition of this class
    */
    setAnew(keyvalArr){
        this.removeAll()
        this._setFromArr(keyvalArr)
    }


    /*
        * Method: invert()
        *   returns a new COMDict instance with inverted key/value pairs
        * Return:
        *   the inverted COMDict instance
    */
    invert(){
        r := new COMDict()
        for k in this.keys()
            Try {
                i := this.item(k)
                r.add(i, k)
            }
            Catch, e
                lol := ""
        return, r
    }


    /*
        * Method: add(key, value)
        *   adds a key value pair but will not update existing on
        * Params:
        *   key:    the key to be add
        *   value:  the value to be associated with the key
    */
    add(key, value){
        this.dict.Add(key, value)
    }


    /*
        * Method: item(key)
        *   returns a specific reference to the given key
        *   equivalent to ahks object[key]
        * Params:
        *   key:    the key to be referenced
        * Return:
        *   a reference to the key value pair
    */
    item(key){
        return, this.dict.Item(key)
    }


    /**
        * Method: __Get(key)
        *   convinience wrapper for the this.item(key) method
        *   for reference, look at its documentation above
        * Params:
        *   key:    the key to be referenced
        * Return:
        *   a reference to the key value pair
        * Note:
        *   In case that any key within your map
        *   (or value in case you invert a map)
        *   is named like a method of this class, it would get cause unwanted sideeffects
    */
    __Get(key){
        return, this.item(key)
    }


    /*
        * Method: remove(key)
        *   removes a specific key from the dictionary
        * Params:
        *   key:    the key to be removed
    */
    remove(key){
        this.dict.Remove(key)
    }


    /*
        * Method: updateKey(key, newKey)
        *   removes a specific key from the dictionary
        * Params:
        *   key:    the key to be updated
        *   newKey: the replacement key to be set
    */
    updateKey(key, newKey){
        this.dict.Key(key) := newKey
    }


    /*
        * Method: exists(key)
        *   checks whether given key exists within the dict
        * Params:
        *   key:    the key to be checked
        * Return:
        *   boolean, true when key exists
    */
    exists(key){
        return, this.dict.Exists(key)
    }


    /*
        * Method: count()
        *   retrieves the number of key/value pairs within the dict
        * Return:
        *   the number of key/value pairs
    */
    count(){
        return, this.dict.count()
    }


    /*
        * Method: items()
        *   retrieves all values present within the dictionary
        * Return:
        *   an standard array of all values
    */
    items(){
        return, this.dict.Items()
    }


    /*
        * Method: keys()
        *   retrieves all keys present within the dictionary
        * Return:
        *   an standard array of all keys
    */
    keys(){
        return, this.dict.Keys()
    }


    /*
        * Method: items()
        *   removes all key/values pairs within the dictionary
        *   clears the hole dict
    */
    removeAll(){
        this.dict.RemoveAll()
    }


    /*
        * Method: _setFromArr(keyvalArr)
        *   adds new key/value pairs from the handed array
        * Params:
        *   keyvalArr:  a dictionary creation formatted array
        *               defined below the definition of this class
    */
    _setFromArr(keyvalArr){
        if(!IsObject(keyvalArr))
            return
        
        for _, o in keyvalArr {
            if(!(IsObject(o)
                    && o.HasKey("key")
                    && o.HasKey("val")))
                continue
            this.add(o.key, o.val)
        }
    }


}


/*
    * The dictionary creation formatted array (short DCFA):
    *   The COMDict class uses a specially formatted standard ahk array
    *   for easy creation of dictionarys, the DCFA.
    *   The DCFA is a normal array containing associative arrays.
    *   These associative arrays have two keys, "key" and "value"
    *   and should be assigned with their values.
    * Example:
        dcfa := []
        dcfa.Push({"key": "a", "val": "ᴀ"}
                , {"key": "b", "val": "ʙ"}
                , {"key": "c", "val": "ᴄ"}
                , {"key": "d", "val": "ᴅ"}
                , {"key": "e", "val": "ᴇ"}
                , {"key": "f", "val": "ꜰ"}
                , {"key": "g", "val": "ɢ"}
                , {"key": "h", "val": "ʜ"}
                , {"key": "i", "val": "ɪ"}
                , {"key": "j", "val": "ᴊ"}
                , {"key": "k", "val": "ᴋ"}
                , {"key": "l", "val": "ʟ"}
                , {"key": "m", "val": "ᴍ"}
                , {"key": "n", "val": "ɴ"}
                , {"key": "o", "val": "ᴏ"}
                , {"key": "p", "val": "ᴘ"}
                , {"key": "q", "val": "ǫ"}
                , {"key": "r", "val": "ʀ"}
                , {"key": "s", "val": "ꜱ"}
                , {"key": "t", "val": "ᴛ"}
                , {"key": "u", "val": "ᴜ"}
                , {"key": "v", "val": "ᴠ"}
                , {"key": "w", "val": "ᴡ"}
                , {"key": "x", "val": "ⅹ"}
                , {"key": "y", "val": "ʏ"}
                , {"key": "z", "val": "ᴢ"})
        dcfa.Push({"key": "A", "val": "ᴀ"}
                , {"key": "B", "val": "ʙ"}
                , {"key": "C", "val": "ᴄ"}
                , {"key": "D", "val": "ᴅ"}
                , {"key": "E", "val": "ᴇ"}
                , {"key": "F", "val": "ꜰ"}
                , {"key": "G", "val": "ɢ"}
                , {"key": "H", "val": "ʜ"}
                , {"key": "I", "val": "ɪ"}
                , {"key": "J", "val": "ᴊ"}
                , {"key": "K", "val": "ᴋ"}
                , {"key": "L", "val": "ʟ"}
                , {"key": "M", "val": "ᴍ"}
                , {"key": "N", "val": "ɴ"}
                , {"key": "O", "val": "ᴏ"}
                , {"key": "P", "val": "ᴘ"}
                , {"key": "Q", "val": "ǫ"}
                , {"key": "R", "val": "ʀ"}
                , {"key": "S", "val": "ꜱ"}
                , {"key": "T", "val": "ᴛ"}
                , {"key": "U", "val": "ᴜ"}
                , {"key": "V", "val": "ᴠ"}
                , {"key": "W", "val": "ᴡ"}
                , {"key": "X", "val": "ⅹ"}
                , {"key": "Y", "val": "ʏ"}
                , {"key": "Z", "val": "ᴢ"})
        SmallCapsDict := new COMDict(dcfa)

        for k in SmallCapsDict.Keys()
        	test .= k " = " SmallCapsDict.item(k) "`n"
        MsgBox, % test
        return
*/