<%@CODEPAGE="65001"%>
<%
    '测试使用 dim 声明和不声明之间的性能差距
    '作者：孙振国
    '创建时间：2022.03.28



    dim numEach     '循环次数
    

    numEach = 1000000

    dim starttime, endtime
    dim starttime2, endtime2
    
    starttime = timer()

    dim i
    for i = 0 to numEach

    next
    endtime = timer()

    starttime2 = timer()
    for r = 0 to numEach

    next
    endtime2 = timer()


    response.Write"dim:" & FormatNumber((endtime-starttime)*1000,3) & "ms"
    response.Write"<br/>"
    response.Write"empty:" & FormatNumber((endtime2-starttime2)*1000,3) & "ms"

%>