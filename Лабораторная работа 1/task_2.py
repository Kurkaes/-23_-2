# TODO Найдите количество книг, которое можно разместить на дискете
ves_simvola = 4
kol_simv_v_1_stroke = 25
kol_strok_na_1_str = 50
kol_str_v_1_knige = 100
V_disketi = 1.44*1024*1024
kol_knig = round(V_disketi / (ves_simvola * kol_simv_v_1_stroke * kol_strok_na_1_str * kol_str_v_1_knige))
print("Количество книг, помещающихся на дискету:", kol_knig)

