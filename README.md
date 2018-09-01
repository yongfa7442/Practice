       Sub TEST2()                             'INSERT
                Columns("A:A").Insert           '在A列左边插入1列
                Dim a%      
                FOR a =2 TO 2 * n + 1 STEP 2
                        Rows(a).Insert          '插入a行
                        a = a + 2
                NEXT
        End Sub
