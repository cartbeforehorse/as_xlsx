PL/SQL Developer Test script 3.0
51
DECLARE
   obj_  json_object_t := json_object_t('{
    "aggregates": {
        "aggregatesCount": 1,
        "aggregateCols": [
            {
                "colId": 4,
                "colName": "Sum of Amount",
                "fn": "sum",
                "value": 2365.42
            }
        ],
        "excelTabHeader": [
            "EUR",
            "GBP",
            "SEK",
            "USD"
        ]
    },
    "direction": "left",
    "count": 54
}');
   agg_obj_ json_object_t := obj_.get_object('aggregates');
   arr_     json_array_t  := obj_.get_object('aggregates').get_array('excelTabHeader');
   agg_arr_ json_array_t  := obj_.get_object('aggregates').get_array('aggregateCols');
   keys_    json_key_list;
   i_obj_   json_object_t;
   val_     NUMBER;
BEGIN
   Dbms_Output.Put_Line ('Object is: ' || obj_.stringify);
   Dbms_Output.Put_Line ('Aggs is: ' || agg_obj_.stringify);
   Dbms_Output.Put_Line ('Array is: ' || arr_.stringify);
   Dbms_Output.Put_Line ('Array 2 is: ' || agg_arr_.stringify);
   keys_ := obj_.get_keys;
   FOR k_ IN keys_.first .. keys_.last LOOP
      Dbms_Output.Put ('k_ => ' || k_);
      Dbms_Output.Put_Line (', keys_(k_) => ' || keys_(k_));
   END LOOP;
   -- Can't do the following line!!
   --   i_obj_ := agg_arr_.get(0);
   val_ := treat (agg_arr_.get(0) as json_object_t).get_number('value');
   Dbms_Output.Put_Line ('Value is: ' || val_);

   -- what happens if you update a value
   Dbms_Output.Put_Line (obj_.stringify);
   --obj_.patch ('direction', 'right'); -- this changes the order of the keys
   obj_.mergepatch ('{"direction":"right"}');
   obj_.patch ('{"direction":"square"}');
   Dbms_Output.Put_Line (obj_.stringify);

END;
0
0
