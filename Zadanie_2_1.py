import win32com.client
import json
import os

nanocad_app = win32com.client.Dispatch("nanoCADx64.Application.23.0")
if nanocad_app is not None:
    ncad_doc = nanocad_app.ActiveDocument
    if ncad_doc is not None:
        for one_layout_index in range(0, ncad_doc.Layouts.Count, 1):
            ncad_Layout = ncad_doc.Layouts.Item(one_layout_index)
            if ncad_Layout.Name == "Model":
                list1 = []
                list2 = []
                list3 = []
                list4 = []
                ncad_Block_for_Layout = ncad_Layout.Block
                for i in ncad_Block_for_Layout:
                    if i.ObjectName == "AcDbBlockReference":
                        if len(list1) == 0:
                            list1.append([i.EffectiveName, 1])
                        else:
                            a = 0
                            b = 0
                            for k, j in enumerate(list1):
                                if list1[k][0] == i.EffectiveName:
                                    a = 1
                                    list1[k][1] += 1
                            if a == 0:
                                list1.append([i.EffectiveName, 1])
                    if i.ObjectName == "AcDbPolyline":
                        if len(list2) == 0:
                            list2.append([i.Layer, i.Length])
                        else:
                            a = 0
                            b = 0
                            for k, j in enumerate(list2):
                                if list2[k][0] == i.Layer:
                                    a = 1
                                    list2[k][1] += i.Length
                            if a == 0:
                                list2.append([i.Layer, i.Length])
                    if i.ObjectName == "AcDbText":
                        if len(list3) == 0:
                            list3.append([i.Layer, len(i.TextString)])
                        else:
                            a = 0
                            b = 0
                            for k, j in enumerate(list3):
                                if list3[k][0] == i.Layer:
                                    a = 1
                                    list3[k][1] += len(i.TextString)
                            if a == 0:
                                list3.append([i.Layer, len(i.TextString)])
                    if i.ObjectName == "AcDbHatch":
                        if len(list4) == 0:
                            list4.append([i.Layer, i.Area])
                        else:
                            a = 0
                            b = 0
                            for k, j in enumerate(list4):
                                if list4[k][0] == i.Layer:
                                    a = 1
                                    list4[k][1] += i.Area
                            if a == 0:
                                list4.append([i.Layer, i.Area])

            if ncad_Layout.Name == "Для вставки таблтиц":
                ncad_Block_for_Layout = ncad_Layout.Block

                def insert_title_as_text(text_to_inserting, numx, numy):
                    center_text = str(numx) + ',' + str(numy) + ',0.0'
                    ncad_Block_for_Layout.AddMText(center_text, 25, text_to_inserting)

                def insert_table(table_to_inserting, numx, numy, length_):
                    center_table = str(numx) + "," + str(numy) + ",0"
                    ncad_Block_for_Layout.AddTable(center_table, len(table_to_inserting)+1, 2, 0.25, length_)


                insert_title_as_text("количество Вхождений блока каждого типа", 6.86, 20.74)
                insert_title_as_text("суммарная длина всех линий с сортировкой по слоям", 22.40, 20.74)
                insert_title_as_text("суммарное количество текстовых символов во всех Однострочных текстах с сортировкой по слоям", 5.10, 9.20)
                insert_title_as_text("суммарная площадь всей штриховки с сортировкой по слоям", 23.69, 9.20)
                insert_table(list1, 4.75, 20.78, 4.1)
                insert_table(list2, 19.75, 20.78, 5.2)
                insert_table(list3, 0.48, 9.24, 9.37)
                insert_table(list4, 20.57, 9.24, 5.75)

                ncad_Block_for_Layout = ncad_Layout.Block
                a = 0
                for i in ncad_Block_for_Layout:
                    if i.ObjectName == "AcDbTable":
                        if a == 3:
                            i.SetTextHeight(7, 0.1)
                            i.SetAlignment(7, 5)
                            for k, j in enumerate(list4):
                                i.SetCellValue(k+1, 0, j[0])
                                i.SetCellValue(k+1, 1, j[1])
                            break
                        elif a == 2:
                            i.SetTextHeight(7, 0.1)
                            i.SetAlignment(7, 5)
                            for k, j in enumerate(list3):
                                i.SetCellValue(k + 1, 0, j[0])
                                i.SetCellValue(k + 1, 1, j[1])
                            a += 1
                        elif a == 1:
                            i.SetTextHeight(7, 0.1)
                            i.SetAlignment(7, 5)
                            for k, j in enumerate(list2):
                                i.SetCellValue(k + 1, 0, j[0])
                                i.SetCellValue(k + 1, 1, j[1])
                            a += 1
                        elif a == 0:
                            i.SetTextHeight(7, 0.1)
                            i.SetAlignment(7, 5)
                            for k, j in enumerate(list1):
                                i.SetCellValue(k + 1, 0, j[0])
                                i.SetCellValue(k + 1, 1, j[1])
                            a += 1

    else:
        print("Doc is not running")
else:
    print("App is not running")
