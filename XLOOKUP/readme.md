=XLOOKUP(lookup_value, lookup_array, return_array, [match_mode], [search_mode])

| Argument      | Description                                                                                                                                                                                                                                                                                                                                                                                                 |
| ------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| lookup\_value | Nilai Acuan                                                                                                                                                                                                                                                                                                                                                                                                 |
| lookup\_array | Kolom Atau Range dimana Nilai Acuan Berada                                                                                                                                                                                                                                                                                                                                                                  |
| return\_array | Kolom Atau Range untuk Nilai yang dicari                                                                                                                                                                                                                                                                                                                                                                    |
| match\_mode   | Ini Optional untuk mode Pencarian.<br>0 - Sama Persis. Jika tidak ditemukan akan #N/A. (default)<br><br>\-1 - Sama persis. Jika tidak ada akan menggunakan Nilai terkecil berikutnya<br><br>1 - Sama Persis. Jika tidak ada akan menggunakan Nilai terbesar berikutnya<br><br>2 - Untuk wildcard match seperti \*, ?, dan ~                                                                                 |
| search\_mode  | Ini Optional Untuk Metode Pencarian<br>1 - Pencarian dimulai dari Items Pertama. (default).<br><br>\-1 - Pencarian dibalik, dimulai dari Item terkahir.<br><br>2 - Metode Pencarian binary di kolom lookup\_array harus berurutan jika tidak akan invalid.<br><br>\-2 - Metode Pencarian binary pada kolom lookup\_array being sorted in descending order. If not sorted, invalid results will be returned. |