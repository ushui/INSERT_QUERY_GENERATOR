Attribute VB_Name = "InsertQueryGeneratorModule"
'''''''''''''''''''''''''''''''''''
''' INSERT QUERY GENERATOR MODULE
''' Version 1.0
'''
''' (C) 2018 ushui
''' Released under the MIT license:
''' http://www.opensource.org/licenses/mit-license.php
'''
''' GitHub: https://github.com/ushui/INSERT_QUERY_GENERATOR
'''''''''''''''''''''''''''''''''''
''' 変更履歴
'''
''' 2018/07/15 Version 1.0
''' 新規作成
'''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''
''' 関数のコメントには、XMLドキュメントコメントを採用しています。
''' 「DocFX」や「Sandcastle」のようなツールを使用してドキュメントを生成するためです。
''' XMLドキュメントコメントについては下記をご覧ください。
'''
''' ドキュメント コメントとして推奨される XML タグ (Visual Basic) | Microsoft Docs
''' https://docs.microsoft.com/ja-jp/dotnet/visual-basic/language-reference/xmldoc/recommended-xml-tags-for-documentation-comments
'''''''''''''''''''''''''''''''''''
Option Explicit

''' <summary>
''' 連想配列から値を取得します。
''' </summary>
''' <param name="dict">連想配列</param>
''' <param name="key">キー</param>
''' <remarks>
''' <para><paramref name="dict"/>と<paramref name="key"/>から値を取得します。</para>
''' </remarks>
''' <returns>
''' <para><paramref name="dict"/>に<paramref name="key"/>があれば値を返し、なければ<c>vbNullString</c>を返します。</para>
''' </returns>
Private Function getItemOfDict(ByRef dict, ByRef key As String) As String

  'キーから値を取得（なければvbNullString）
  If Not dict.Exists(key) Then
    getItemOfDict = vbNullString
  Else
    getItemOfDict = dict.Item(key)
  End If

End Function

''' <summary>
''' INSERT_xxxxの引数検査とエラーメッセージ取得
''' </summary>
''' <param name="tableName">テーブル名</param>
''' <param name="types">データ型一覧</param>
''' <param name="clmns">カラム一覧</param>
''' <param name="values">データ一覧</param>
''' <param name="lineFeed">改行文字</param>
''' <param name="toReplaceNull">NULLとして扱う文字列</param>
''' <remarks>
''' <para>INSERT_xxxxの先頭で呼びます。</para>
''' <para>それらに指定した引数が正しいか否かをチェックし、検査結果に応じてメッセージを返します。</para>
''' <para>引数が誤りと判断するケースは下記です。</para>
''' <list type="bullet">
''' <item><description>指定したセルの数が同じでない場合</description></item>
''' <item><description>誤った改行文字を指定した場合</description></item>
''' </list>
''' </remarks>
''' <returns>
''' <para>引数が正しければ<c>vbNullString</c>、誤っていればエラーメッセージを返します。</para>
''' </returns>
Private Function getMsgIfIncorrectArgs(ByRef tableName As String, _
                                       ByRef types As Range, _
                                       ByRef clmns As Range, _
                                       ByRef values As Range, _
                                       ByRef lineFeed As String, _
                                       ByRef toReplaceNull As String) As String

  '指定したセルの数が同じでない場合
  If types.Count <> clmns.Count Or clmns.Count <> values.Count Then
    getMsgIfIncorrectArgs = "ARGUMENTS ERROR: The number of data types, columns, data must match."
    Exit Function
  End If
  '誤った改行文字を指定した場合
  If StrPtr(getInsertableLineFeedCodeOrcl(lineFeed)) = 0 Then
    getMsgIfIncorrectArgs = "ARGUMENTS ERROR: Please specify either 'CRLF' 'CR' 'LF' for the line feed code."
    Exit Function
  End If

  getMsgIfIncorrectArgs = vbNullString

End Function

''' <summary>
''' フォーマットグループ生成と取得（Oracle Database用）
''' </summary>
''' <remarks>
''' <para>データ型ごとに定義されたフォーマットグループの連想配列を生成し、取得します。</para>
''' </remarks>
''' <returns>
''' <para>データ型ごとのフォーマットグループ</para>
''' </returns>
Private Function getDictOfFormatGroupByDataTypeOrcl() As Object

  Dim formatGroupByDataType As Object: Set formatGroupByDataType = CreateObject("Scripting.Dictionary")
  With formatGroupByDataType
    'キーの大文字・小文字は区別しない
    .CompareMode = vbTextCompare

    '文字列型
    .Add "CHAR", "CHAR"
    .Add "LONG", "CHAR"
    .Add "NCHAR", "CHAR"
    .Add "NVARCHAR2", "CHAR"
    .Add "VARCHAR2", "CHAR"
    '数値型
    .Add "NUMBER", "NUMBER"
    .Add "BINARY_FLOAT", "NUMBER"
    .Add "BINARY_DOUBLE", "NUMBER"
    '日付型
    .Add "DATE", "DATE"
    '時刻型
    .Add "TIMESTAMP", "TIMESTAMP"
    'バイナリ型・ラージオブジェクト型
    .Add "RAW", "RAW"
    .Add "LONG RAW", "RAW"
    .Add "BLOB", "RAW"
    .Add "CLOB", "RAW"
    .Add "NCLOB", "RAW"
    'バイナリ型（BFILE）
    .Add "BFILE", "BFILE"
  End With

  Set getDictOfFormatGroupByDataTypeOrcl = formatGroupByDataType

End Function

''' <summary>
''' フォーマットグループ名取得（Oracle Database用）
''' </summary>
''' <param name="dataType">データ型</param>
''' <remarks>
''' <para><paramref name="dataType"/>に対するフォーマットグループ名を取得します。</para>
''' </remarks>
''' <returns>
''' <para><paramref name="dataType"/>がフォーマットグループに属していればフォーマットグループ名、属していなければ<c>vbNullString</c>を返します。</para>
''' </returns>
Private Function getFormatGroupNameOrcl(ByRef dataType As String) As String

  Static dictFormatGroupByDataType As Object
  '初期化されている場合のみフォーマットグループを取得（Static変数であることに留意）
  If dictFormatGroupByDataType Is Nothing Then
    Set dictFormatGroupByDataType = getDictOfFormatGroupByDataTypeOrcl()
  End If

  getFormatGroupNameOrcl = getItemOfDict(dictFormatGroupByDataType, dataType)

End Function

''' <summary>
''' INSERT可能な改行コードを取得（Oracle Database用）
''' </summary>
''' <param name="lineFeed">改行文字</param>
''' <remarks>
''' <para><paramref name="lineFeed"/>に対応する、INSERT可能な改行コードを取得します（Oracle Database用）。</para>
''' </remarks>
''' <returns>
''' <para>引数として与えた<paramref name="lineFeed"/>によって、下記の値を返します。いずれにも当てはまらなかった場合は<c>vbNullString</c>を返します。</para>
''' <list type="bullet">
''' <item><description><c>CRLF</c>の場合: <c>' || CHR(13) || CHR(10) || '」</c></description></item>
''' <item><description><c>CR</c>の場合: <c>' || CHR(13) || '</c></description></item>
''' <item><description><c>LF</c>の場合: <c>' || CHR(10) || '</c></description></item>
''' </list>
''' </returns>
Private Function getInsertableLineFeedCodeOrcl(ByRef lineFeed As String) As String

  Select Case lineFeed
    Case "CRLF"
      getInsertableLineFeedCodeOrcl = "' || CHR(13) || CHR(10) || '"
    Case "LF"
      getInsertableLineFeedCodeOrcl = "' || CHR(10) || '"
    Case "CR"
      getInsertableLineFeedCodeOrcl = "' || CHR(13) || '"
    Case Else
      getInsertableLineFeedCodeOrcl = vbNullString
  End Select

End Function

''' <summary>
''' INSERT文を生成（Oracle Database用）
''' </summary>
''' <param name="tableName">テーブル名</param>
''' <param name="types">データ型一覧</param>
''' <param name="clmns">カラム一覧</param>
''' <param name="values">データ一覧</param>
''' <param name="lineFeed">改行文字</param>
''' <param name="toReplaceNull">NULLとして扱う文字列</param>
''' <remarks>
''' <para>ユーザー定義関数です。</para>
''' <para>指定した引数からINSERT文を生成します（Oracle Database用）。</para>
''' <para></para>
''' <para>●引数について</para>
''' <para>生成を行う前に引数の検査を行い、誤っていればエラーメッセージを生成します。</para>
''' <para>また、<paramref name="toReplaceNull"/>を省略した場合、空文字をNULLとして扱います。</para>
''' <para></para>
''' <para>●変換に対応しているデータ型について</para>
''' <para>期間データ型とBFILE型を除くOracle組込みデータ型に対応しております。</para>
''' <para>XMLTYPE型等には対応しておりませんが、セルへ<c>xmltype('<?xml version="1.0"?><Test></Test>')</c>のように直接入力することでINSERTは可能です。</para>
''' <para></para>
''' <para>●変換仕様</para>
''' <para>VALUES句の引数を生成する際は、事前に<paramref name="values"/>の変換を行い、<paramref name="types"/>ごとの書式に合わせた文字列を生成します。</para>
''' <para>○文字列型</para>
''' <para>INSERT可能な文字列へ置換し、<c>'</c>で括ります。</para>
''' <para>○数値型</para>
''' <para>変換しません。</para>
''' <para>○日付型</para>
''' <para>TO_DATE関数を生成します。</para>
''' <para>データ型には<c>:</c>を付けることによって書式を設定することができます。これは省略可能です。</para>
''' <para>・第一引数：データ</para>
''' <para>・第二引数：<c>:</c>がない場合は省略する。<c>:</c>より先に指定した文字列を<c>'</c>で括った文字列</para>
''' <para>○時刻型</para>
''' <para>TO_TIMESTAMP関数を生成します。</para>
''' <para>データ型には<c>:</c>を付けることによって書式を設定することができます。これは省略可能です。</para>
''' <para>・第一引数：データ</para>
''' <para>・第二引数：<c>:</c>がない場合は省略する。<c>:</c>より先に指定した文字列を<c>'</c>で括った文字列</para>
''' <para>○バイナリ型・ラージオブジェクト型（BFILE型除く）</para>
''' <para>HEXTORAW関数を生成します。</para>
''' <para>・第一引数：データ</para>
''' <para>○その他</para>
''' <para>変換しません。</para>
''' <para></para>
''' <para>●注意事項</para>
''' <para>Oracle Databaseでは長さ0の文字列をNULLとして扱います。</para>
''' <para></para>
''' <para>●リファレンス</para>
''' <para>https://docs.oracle.com/cd/E57425_01/121/SQLRF/sql_elements003.htm</para>
''' <para></para>
''' <para>●その他/para>
''' <para>Static関数であるため、変数の初期化と再利用には注意を払ってください。</para>
''' <para>これは本関数が幾度も呼ばれることを想定しており、関数内の変数の領域をその都度で確保せずに済むようにしているためです。</para>
''' </remarks>
''' <returns>
''' INSERT文
''' </returns>
Public Static Function INSERT_ORCL(tableName As String, _
                                   types As Range, _
                                   clmns As Range, _
                                   values As Range, _
                                   lineFeed As String, _
                                   Optional toReplaceNull As String = "") As String

  ' --------------------
  ' 引数エラーチェック
  ' --------------------
  Dim errMsg As String: errMsg = getMsgIfIncorrectArgs(tableName, types, clmns, values, lineFeed, toReplaceNull)
  'エラーメッセージがvbNullString以外である場合
  If StrPtr(errMsg) <> 0 Then
    'MsgBoxは連続で表示されてしまう可能性が高いため使用しない
    INSERT_ORCL = errMsg
    Exit Function
  End If

  ' --------------------
  ' 定義部
  ' --------------------
  '改行コード
  Dim insertableLineFeedCode As String: insertableLineFeedCode = getInsertableLineFeedCodeOrcl(lineFeed)
  '文字数
  Dim lenDate As Long: lenDate = Len("DATE")
  Dim lenTimeStamp As Long: lenTimeStamp = Len("TIMESTAMP")
  'カラム、データ格納用（添え字は1から）
  Dim arrClmns As Variant, arrValues As Variant
  ReDim arrClmns(1 To clmns.Count), arrValues(1 To values.Count)
  'ループ用
  Dim i As Long
  'その他
  Dim tmpReplace As String
  Dim tmpIdxAtHyphen As Long
  Dim tmpArrIntervalsStr() As String

  ' --------------------
  ' 処理部
  ' --------------------
  'カラム一覧を配列化
  For i = 1 To clmns.Count
    arrClmns(i) = CStr(clmns.Item(i).Value)
  Next

  'データ一覧を配列化し、データ型ごとの書式に合わせた文字列を生成
  For i = 1 To values.Count
    'データがNULL
    If values.Item(i).Value = toReplaceNull Then
      arrValues(i) = "NULL"
    Else
      '文字列型
      If getFormatGroupNameOrcl(types.Item(i).Value) = "CHAR" Then
        'エスケープ等
        tmpReplace = values.Item(i).Value
        tmpReplace = Replace(tmpReplace, "'", "''")
        tmpReplace = Replace(tmpReplace, vbTab, "' || CHR(9) || '")
        tmpReplace = Replace(tmpReplace, vbLf, insertableLineFeedCode)
        'シングルクォートで括る
        arrValues(i) = Join(Array("'", tmpReplace, "'"), "")

      '数値型
      ElseIf getFormatGroupNameOrcl(types.Item(i).Value) = "NUMBER" Then
        arrValues(i) = CStr(values.Item(i).Value)

      '日付型
      ElseIf getFormatGroupNameOrcl(Left(types.Item(i).Value, lenDate)) = "DATE" Then
        If Mid(types.Item(i).Value, lenDate + 1, 1) = ":" Then
          'TO_DATE('データ')
          arrValues(i) = Join(Array("TO_DATE('", values.Item(i).Value, "','", Mid(types.Item(i).Value, lenDate + 2), "')"), "")
        Else
          'TO_DATE('データ', '書式')
          arrValues(i) = Join(Array("TO_DATE('", values.Item(i).Value, "')"), "")
        End If

      '時刻型
      ElseIf getFormatGroupNameOrcl(Left(types.Item(i).Value, lenTimeStamp)) = "TIMESTAMP" Then
        If Mid(types.Item(i).Value, lenTimeStamp + 1, 1) = ":" Then
          'TO_TIMESTAMP('データ')
          arrValues(i) = Join(Array("TO_TIMESTAMP('", values.Item(i).Value, "','", Mid(types.Item(i).Value, lenTimeStamp + 2), "')"), "")
        Else
          'TO_TIMESTAMP('データ', '書式')
          arrValues(i) = Join(Array("TO_DATE('", values.Item(i).Value, "')"), "")
        End If

      'バイナリ型・ラージオブジェクト型
      ElseIf getFormatGroupNameOrcl(Left(types.Item(i).Value, lenTimeStamp)) = "RAW" Then
        'DECODE('データ', 'HEX')
        arrValues(i) = Join(Array("HEXTORAW('", values.Item(i).Value, "')"), "")

      '上記に当てはまらないデータ型
      Else
        arrValues(i) = CStr(values.Item(i).Value)

      End If
    End If
  Next

  'INSERT文を生成
  INSERT_ORCL = Join(Array("INSERT INTO ", tableName, "(", Join(arrClmns, ","), ") VALUES(", Join(arrValues, ","), ");"), "")

End Function
