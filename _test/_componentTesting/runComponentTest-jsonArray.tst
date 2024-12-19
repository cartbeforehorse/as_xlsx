PL/SQL Developer Test script 3.0
142
DECLARE
   arr1_    json_array_t := json_array_t(); -- must be initiated
   arr_     json_array_t := json_array_t('["one", "two", "three", 5]');
   arr_e_   json_array_t := json_array_t('[]');
   arr_f_   json_array_t := json_array_t();
   arr_g_   json_array_t := json_array_t();
   arr_arr_ json_array_t := json_array_t('[[]]');
   arr_x_   json_array_t := json_array_t('["a", "b"]');
   arr_y_   json_array_t := json_array_t('["c", "d"]');
   arr_z_   json_array_t := json_array_t('["a", ["b", "c"]]');
   val_     VARCHAR2(50);
BEGIN
   arr1_.append_all(arr_);
   Dbms_Output.Put_Line ('First arr: ' || arr1_.stringify);
   arr_.put (2, 'four');
   arr_.put (1, 'five');
   arr_.put (0, 'six');
   arr_.put (3, 'three override', true);
   Dbms_Output.Put_Line ('The array =====>');
   FOR i_ IN 0 .. arr_.get_size - 1 LOOP
      IF arr_.get(i_).is_string THEN
         Dbms_Output.Put_Line (i_ || ': ' || arr_.get_string(i_));
      ELSE
         Dbms_Output.Put_Line ('nr: ' || arr_.get_number(i_));
      END IF;
   END LOOP;
   Dbms_Output.Put_Line ('<========');
   Dbms_Output.Put_Line ('arr_e_: ' || arr_e_.stringify || ', size: ' || arr_e_.get_size);
   IF arr_e_.is_null THEN
      Dbms_Output.Put_Line ('arr_e_ is null by virtue of being empty');
   ELSE
      Dbms_Output.Put_Line ('arr_e_ is not null, even though it''s empty');
   END IF;
   IF arr_f_.is_null THEN
      Dbms_Output.Put_Line ('arr_f_ is null by virtue of being empty');
   ELSE
      Dbms_Output.Put_Line ('arr_f_ is not null, even though it''s empty');
   END IF;
   arr_f_.append('chicken');
   arr_f_.append('');
   Dbms_Output.Put_Line ('arr_f_ with an appended chicken => ' || arr_f_.stringify);
   Dbms_Output.Put_Line ('arr_g_ => ' || arr_g_.stringify);
   IF arr_arr_.get_size != 0 THEN
      Dbms_Output.Put_Line ('Length of array is: ' || arr_arr_.get_size);
      Dbms_Output.Put_Line ('Array element zero: ' || arr_arr_.get(0).stringify);
      Dbms_Output.Put_Line ('Array element is array: ' || to_char(arr_arr_.get(0).is_array));
   END IF;
   arr_x_.append_all (arr_y_);
   arr_x_.append ('g');
   Dbms_Output.Put_Line ('Array x: ' || arr_x_.stringify);
   Dbms_Output.Put_Line ('Array x: ' || arr_x_.get_type(1));
   FOR ix_ IN 0 .. arr_z_.get_size-1 LOOP
      Dbms_Output.Put_Line ('Element type: ' || arr_z_.get_type(ix_));
      IF arr_z_.get(ix_).is_array THEN
         Dbms_Output.Put_Line (
            'Element ' || ix_ || ': is an array => ' || treat(arr_z_.get(ix_) as json_array_t).stringify
         );
         Dbms_Output.Put_Line ('Inner array size: ' || treat (arr_z_.get(1) as json_array_t).get_size);
      ELSE
         Dbms_Output.Put_Line ('Element ' || ix_ || ': is NOT an array');
      END IF;
   END LOOP;
   -- The following cannot be done:
   --treat (arr_z_.get(1) as json_array_t).append('d');
   --Dbms_Output.Put_Line ('New Z array: ' || arr_z_.stringify);

   -- what happens with nulls?
   arr_z_ := json_array_t();
   --arr_z_.append (null); -- cannot be null
   arr_z_.append (val_);
   val_ := '';
   arr_z_.append (val_);
   arr_z_.append ('1');
   arr_z_.append ('');
   Dbms_Output.Put_Line ('Testing null and emptry-string values: ' || arr_z_.stringify);
   FOR ix_ IN 0 .. arr_z_.get_size-1 LOOP
      Dbms_Output.Put_Line (CASE WHEN arr_z_.get_string(ix_) IS NOT null THEN arr_z_.get_string(ix_) ELSE '<< null >>' END);
   END LOOP;

   -- Adding up null values
   DECLARE
      arr_nr_   json_array_t := json_array_t('[6, "", null, 7]');
      ix_nr_    NUMBER       := 0;
      sum_      NUMBER       := 0;
      have_rec_ BOOLEAN      := false;
   BEGIN
      Dbms_Output.Put_Line ('');
      Dbms_Output.Put_Line ('Adding null values: ' || arr_nr_.stringify);
      FOR ix_ IN 0 .. arr_nr_.get_size-1 LOOP
         ix_nr_ := arr_nr_.get_number(ix_);
         sum_ := sum_ + nvl (ix_nr_,0);
         IF sum_ IS NOT null THEN
            have_rec_ := true;
         END IF;
         Dbms_Output.Put_Line ('Number is: ' || nvl(to_char(ix_nr_),'<<null>>'));
         Dbms_Output.Put_Line ('Sum is: ' || sum_);
      END LOOP;
      Dbms_Output.Put_Line (CASE WHEN have_rec_ THEN 'we have a record' ELSE 'we do not have a record' END);
   END;
   -- accessing array objeccts that aren't there
   DECLARE
      arr_  json_array_t := json_array_t('["a", "b", "c", "d", "e", ""]');
      val_  VARCHAR2(20);
      el_   json_element_t;
   BEGIN
      val_ := arr_.get_string(2);
      Dbms_Output.Put_Line ('location 2 holds: ' || val_);
      val_ := arr_.get_string(20);
      Dbms_Output.Put_Line ('location 20 holds: ' || val_);
      IF arr_.has_value('f') THEN
         Dbms_Output.Put_Line ('"f" is in array');
      ELSE
         Dbms_Output.Put_Line ('"f" not in array');
      END IF;
      IF arr_.has_value('c') THEN
         Dbms_Output.Put_Line ('"c" is in array');
      ELSE
         Dbms_Output.Put_Line ('"c" not in array');
      END IF;
      IF arr_.has_value('') THEN
         Dbms_Output.Put_Line ('null is in array');
      ELSE
         Dbms_Output.Put_Line ('no null in this array!');
      END IF;
      arr_.put_null (5, true);
      IF arr_.has_value(to_number(null)) THEN
         Dbms_Output.Put_Line ('null is in array');
      ELSE
         Dbms_Output.Put_Line ('no null in this array!');
      END IF;
      Dbms_Output.Put_Line ('New arr: ' || arr_.stringify);
   END;

   DECLARE
      arr_  json_array_t := json_array_t('[]'); -- same as ()
      arr2_ json_array_t := json_array_t('[1, 2, 3, 4]');
   BEGIN
      Dbms_Output.Put_Line ('TEST 4 ==>>');
      arr_.append('a'); --.append_all(arr2_);
      Dbms_Output.Put_Line (arr_.stringify);
   END;
END;
0
0
