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
''' �ύX����
'''
''' 2018/07/15 Version 1.0
''' �V�K�쐬
'''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''
''' �֐��̃R�����g�ɂ́AXML�h�L�������g�R�����g���̗p���Ă��܂��B
''' �uDocFX�v��uSandcastle�v�̂悤�ȃc�[�����g�p���ăh�L�������g�𐶐����邽�߂ł��B
''' XML�h�L�������g�R�����g�ɂ��Ă͉��L���������������B
'''
''' �h�L�������g �R�����g�Ƃ��Đ�������� XML �^�O (Visual Basic) | Microsoft Docs
''' https://docs.microsoft.com/ja-jp/dotnet/visual-basic/language-reference/xmldoc/recommended-xml-tags-for-documentation-comments
'''''''''''''''''''''''''''''''''''
Option Explicit

''' <summary>
''' �A�z�z�񂩂�l���擾���܂��B
''' </summary>
''' <param name="dict">�A�z�z��</param>
''' <param name="key">�L�[</param>
''' <remarks>
''' <para><paramref name="dict"/>��<paramref name="key"/>����l���擾���܂��B</para>
''' </remarks>
''' <returns>
''' <para><paramref name="dict"/>��<paramref name="key"/>������Βl��Ԃ��A�Ȃ����<c>vbNullString</c>��Ԃ��܂��B</para>
''' </returns>
Private Function getItemOfDict(ByRef dict, ByRef key As String) As String

  '�L�[����l���擾�i�Ȃ����vbNullString�j
  If Not dict.Exists(key) Then
    getItemOfDict = vbNullString
  Else
    getItemOfDict = dict.Item(key)
  End If

End Function

''' <summary>
''' INSERT_xxxx�̈��������ƃG���[���b�Z�[�W�擾
''' </summary>
''' <param name="tableName">�e�[�u����</param>
''' <param name="types">�f�[�^�^�ꗗ</param>
''' <param name="clmns">�J�����ꗗ</param>
''' <param name="values">�f�[�^�ꗗ</param>
''' <param name="lineFeed">���s����</param>
''' <param name="toReplaceNull">NULL�Ƃ��Ĉ���������</param>
''' <remarks>
''' <para>INSERT_xxxx�̐擪�ŌĂт܂��B</para>
''' <para>�����Ɏw�肵�����������������ۂ����`�F�b�N���A�������ʂɉ����ă��b�Z�[�W��Ԃ��܂��B</para>
''' <para>���������Ɣ��f����P�[�X�͉��L�ł��B</para>
''' <list type="bullet">
''' <item><description>�w�肵���Z���̐��������łȂ��ꍇ</description></item>
''' <item><description>��������s�������w�肵���ꍇ</description></item>
''' </list>
''' </remarks>
''' <returns>
''' <para>���������������<c>vbNullString</c>�A����Ă���΃G���[���b�Z�[�W��Ԃ��܂��B</para>
''' </returns>
Private Function getMsgIfIncorrectArgs(ByRef tableName As String, _
                                       ByRef types As Range, _
                                       ByRef clmns As Range, _
                                       ByRef values As Range, _
                                       ByRef lineFeed As String, _
                                       ByRef toReplaceNull As String) As String

  '�w�肵���Z���̐��������łȂ��ꍇ
  If types.Count <> clmns.Count Or clmns.Count <> values.Count Then
    getMsgIfIncorrectArgs = "ARGUMENTS ERROR: The number of data types, columns, data must match."
    Exit Function
  End If
  '��������s�������w�肵���ꍇ
  If StrPtr(getInsertableLineFeedCodeOrcl(lineFeed)) = 0 Then
    getMsgIfIncorrectArgs = "ARGUMENTS ERROR: Please specify either 'CRLF' 'CR' 'LF' for the line feed code."
    Exit Function
  End If

  getMsgIfIncorrectArgs = vbNullString

End Function

''' <summary>
''' �t�H�[�}�b�g�O���[�v�����Ǝ擾�iOracle Database�p�j
''' </summary>
''' <remarks>
''' <para>�f�[�^�^���Ƃɒ�`���ꂽ�t�H�[�}�b�g�O���[�v�̘A�z�z��𐶐����A�擾���܂��B</para>
''' </remarks>
''' <returns>
''' <para>�f�[�^�^���Ƃ̃t�H�[�}�b�g�O���[�v</para>
''' </returns>
Private Function getDictOfFormatGroupByDataTypeOrcl() As Object

  Dim formatGroupByDataType As Object: Set formatGroupByDataType = CreateObject("Scripting.Dictionary")
  With formatGroupByDataType
    '�L�[�̑啶���E�������͋�ʂ��Ȃ�
    .CompareMode = vbTextCompare

    '������^
    .Add "CHAR", "CHAR"
    .Add "LONG", "CHAR"
    .Add "NCHAR", "CHAR"
    .Add "NVARCHAR2", "CHAR"
    .Add "VARCHAR2", "CHAR"
    '���l�^
    .Add "NUMBER", "NUMBER"
    .Add "BINARY_FLOAT", "NUMBER"
    .Add "BINARY_DOUBLE", "NUMBER"
    '���t�^
    .Add "DATE", "DATE"
    '�����^
    .Add "TIMESTAMP", "TIMESTAMP"
    '�o�C�i���^�E���[�W�I�u�W�F�N�g�^
    .Add "RAW", "RAW"
    .Add "LONG RAW", "RAW"
    .Add "BLOB", "RAW"
    .Add "CLOB", "RAW"
    .Add "NCLOB", "RAW"
    '�o�C�i���^�iBFILE�j
    .Add "BFILE", "BFILE"
  End With

  Set getDictOfFormatGroupByDataTypeOrcl = formatGroupByDataType

End Function

''' <summary>
''' �t�H�[�}�b�g�O���[�v���擾�iOracle Database�p�j
''' </summary>
''' <param name="dataType">�f�[�^�^</param>
''' <remarks>
''' <para><paramref name="dataType"/>�ɑ΂���t�H�[�}�b�g�O���[�v�����擾���܂��B</para>
''' </remarks>
''' <returns>
''' <para><paramref name="dataType"/>���t�H�[�}�b�g�O���[�v�ɑ����Ă���΃t�H�[�}�b�g�O���[�v���A�����Ă��Ȃ����<c>vbNullString</c>��Ԃ��܂��B</para>
''' </returns>
Private Function getFormatGroupNameOrcl(ByRef dataType As String) As String

  Static dictFormatGroupByDataType As Object
  '����������Ă���ꍇ�̂݃t�H�[�}�b�g�O���[�v���擾�iStatic�ϐ��ł��邱�Ƃɗ��Ӂj
  If dictFormatGroupByDataType Is Nothing Then
    Set dictFormatGroupByDataType = getDictOfFormatGroupByDataTypeOrcl()
  End If

  getFormatGroupNameOrcl = getItemOfDict(dictFormatGroupByDataType, dataType)

End Function

''' <summary>
''' INSERT�\�ȉ��s�R�[�h���擾�iOracle Database�p�j
''' </summary>
''' <param name="lineFeed">���s����</param>
''' <remarks>
''' <para><paramref name="lineFeed"/>�ɑΉ�����AINSERT�\�ȉ��s�R�[�h���擾���܂��iOracle Database�p�j�B</para>
''' </remarks>
''' <returns>
''' <para>�����Ƃ��ė^����<paramref name="lineFeed"/>�ɂ���āA���L�̒l��Ԃ��܂��B������ɂ����Ă͂܂�Ȃ������ꍇ��<c>vbNullString</c>��Ԃ��܂��B</para>
''' <list type="bullet">
''' <item><description><c>CRLF</c>�̏ꍇ: <c>' || CHR(13) || CHR(10) || '�v</c></description></item>
''' <item><description><c>CR</c>�̏ꍇ: <c>' || CHR(13) || '</c></description></item>
''' <item><description><c>LF</c>�̏ꍇ: <c>' || CHR(10) || '</c></description></item>
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
''' INSERT���𐶐��iOracle Database�p�j
''' </summary>
''' <param name="tableName">�e�[�u����</param>
''' <param name="types">�f�[�^�^�ꗗ</param>
''' <param name="clmns">�J�����ꗗ</param>
''' <param name="values">�f�[�^�ꗗ</param>
''' <param name="lineFeed">���s����</param>
''' <param name="toReplaceNull">NULL�Ƃ��Ĉ���������</param>
''' <remarks>
''' <para>���[�U�[��`�֐��ł��B</para>
''' <para>�w�肵����������INSERT���𐶐����܂��iOracle Database�p�j�B</para>
''' <para></para>
''' <para>�������ɂ���</para>
''' <para>�������s���O�Ɉ����̌������s���A����Ă���΃G���[���b�Z�[�W�𐶐����܂��B</para>
''' <para>�܂��A<paramref name="toReplaceNull"/>���ȗ������ꍇ�A�󕶎���NULL�Ƃ��Ĉ����܂��B</para>
''' <para></para>
''' <para>���ϊ��ɑΉ����Ă���f�[�^�^�ɂ���</para>
''' <para>���ԃf�[�^�^��BFILE�^������Oracle�g���݃f�[�^�^�ɑΉ����Ă���܂��B</para>
''' <para>XMLTYPE�^���ɂ͑Ή����Ă���܂��񂪁A�Z����<c>xmltype('<?xml version="1.0"?><Test></Test>')</c>�̂悤�ɒ��ړ��͂��邱�Ƃ�INSERT�͉\�ł��B</para>
''' <para></para>
''' <para>���ϊ��d�l</para>
''' <para>VALUES��̈����𐶐�����ۂ́A���O��<paramref name="values"/>�̕ϊ����s���A<paramref name="types"/>���Ƃ̏����ɍ��킹��������𐶐����܂��B</para>
''' <para>��������^</para>
''' <para>INSERT�\�ȕ�����֒u�����A<c>'</c>�Ŋ���܂��B</para>
''' <para>�����l�^</para>
''' <para>�ϊ����܂���B</para>
''' <para>�����t�^</para>
''' <para>TO_DATE�֐��𐶐����܂��B</para>
''' <para>�f�[�^�^�ɂ�<c>:</c>��t���邱�Ƃɂ���ď�����ݒ肷�邱�Ƃ��ł��܂��B����͏ȗ��\�ł��B</para>
''' <para>�E�������F�f�[�^</para>
''' <para>�E�������F<c>:</c>���Ȃ��ꍇ�͏ȗ�����B<c>:</c>����Ɏw�肵���������<c>'</c>�Ŋ�����������</para>
''' <para>�������^</para>
''' <para>TO_TIMESTAMP�֐��𐶐����܂��B</para>
''' <para>�f�[�^�^�ɂ�<c>:</c>��t���邱�Ƃɂ���ď�����ݒ肷�邱�Ƃ��ł��܂��B����͏ȗ��\�ł��B</para>
''' <para>�E�������F�f�[�^</para>
''' <para>�E�������F<c>:</c>���Ȃ��ꍇ�͏ȗ�����B<c>:</c>����Ɏw�肵���������<c>'</c>�Ŋ�����������</para>
''' <para>���o�C�i���^�E���[�W�I�u�W�F�N�g�^�iBFILE�^�����j</para>
''' <para>HEXTORAW�֐��𐶐����܂��B</para>
''' <para>�E�������F�f�[�^</para>
''' <para>�����̑�</para>
''' <para>�ϊ����܂���B</para>
''' <para></para>
''' <para>�����ӎ���</para>
''' <para>Oracle Database�ł͒���0�̕������NULL�Ƃ��Ĉ����܂��B</para>
''' <para></para>
''' <para>�����t�@�����X</para>
''' <para>https://docs.oracle.com/cd/E57425_01/121/SQLRF/sql_elements003.htm</para>
''' <para></para>
''' <para>�����̑�/para>
''' <para>Static�֐��ł��邽�߁A�ϐ��̏������ƍė��p�ɂ͒��ӂ𕥂��Ă��������B</para>
''' <para>����͖{�֐������x���Ă΂�邱�Ƃ�z�肵�Ă���A�֐����̕ϐ��̗̈�����̓s�x�Ŋm�ۂ����ɍςނ悤�ɂ��Ă��邽�߂ł��B</para>
''' </remarks>
''' <returns>
''' INSERT��
''' </returns>
Public Static Function INSERT_ORCL(tableName As String, _
                                   types As Range, _
                                   clmns As Range, _
                                   values As Range, _
                                   lineFeed As String, _
                                   Optional toReplaceNull As String = "") As String

  ' --------------------
  ' �����G���[�`�F�b�N
  ' --------------------
  Dim errMsg As String: errMsg = getMsgIfIncorrectArgs(tableName, types, clmns, values, lineFeed, toReplaceNull)
  '�G���[���b�Z�[�W��vbNullString�ȊO�ł���ꍇ
  If StrPtr(errMsg) <> 0 Then
    'MsgBox�͘A���ŕ\������Ă��܂��\�����������ߎg�p���Ȃ�
    INSERT_ORCL = errMsg
    Exit Function
  End If

  ' --------------------
  ' ��`��
  ' --------------------
  '���s�R�[�h
  Dim insertableLineFeedCode As String: insertableLineFeedCode = getInsertableLineFeedCodeOrcl(lineFeed)
  '������
  Dim lenDate As Long: lenDate = Len("DATE")
  Dim lenTimeStamp As Long: lenTimeStamp = Len("TIMESTAMP")
  '�J�����A�f�[�^�i�[�p�i�Y������1����j
  Dim arrClmns As Variant, arrValues As Variant
  ReDim arrClmns(1 To clmns.Count), arrValues(1 To values.Count)
  '���[�v�p
  Dim i As Long
  '���̑�
  Dim tmpReplace As String
  Dim tmpIdxAtHyphen As Long
  Dim tmpArrIntervalsStr() As String

  ' --------------------
  ' ������
  ' --------------------
  '�J�����ꗗ��z��
  For i = 1 To clmns.Count
    arrClmns(i) = CStr(clmns.Item(i).Value)
  Next

  '�f�[�^�ꗗ��z�񉻂��A�f�[�^�^���Ƃ̏����ɍ��킹��������𐶐�
  For i = 1 To values.Count
    '�f�[�^��NULL
    If values.Item(i).Value = toReplaceNull Then
      arrValues(i) = "NULL"
    Else
      '������^
      If getFormatGroupNameOrcl(types.Item(i).Value) = "CHAR" Then
        '�G�X�P�[�v��
        tmpReplace = values.Item(i).Value
        tmpReplace = Replace(tmpReplace, "'", "''")
        tmpReplace = Replace(tmpReplace, vbTab, "' || CHR(9) || '")
        tmpReplace = Replace(tmpReplace, vbLf, insertableLineFeedCode)
        '�V���O���N�H�[�g�Ŋ���
        arrValues(i) = Join(Array("'", tmpReplace, "'"), "")

      '���l�^
      ElseIf getFormatGroupNameOrcl(types.Item(i).Value) = "NUMBER" Then
        arrValues(i) = CStr(values.Item(i).Value)

      '���t�^
      ElseIf getFormatGroupNameOrcl(Left(types.Item(i).Value, lenDate)) = "DATE" Then
        If Mid(types.Item(i).Value, lenDate + 1, 1) = ":" Then
          'TO_DATE('�f�[�^')
          arrValues(i) = Join(Array("TO_DATE('", values.Item(i).Value, "','", Mid(types.Item(i).Value, lenDate + 2), "')"), "")
        Else
          'TO_DATE('�f�[�^', '����')
          arrValues(i) = Join(Array("TO_DATE('", values.Item(i).Value, "')"), "")
        End If

      '�����^
      ElseIf getFormatGroupNameOrcl(Left(types.Item(i).Value, lenTimeStamp)) = "TIMESTAMP" Then
        If Mid(types.Item(i).Value, lenTimeStamp + 1, 1) = ":" Then
          'TO_TIMESTAMP('�f�[�^')
          arrValues(i) = Join(Array("TO_TIMESTAMP('", values.Item(i).Value, "','", Mid(types.Item(i).Value, lenTimeStamp + 2), "')"), "")
        Else
          'TO_TIMESTAMP('�f�[�^', '����')
          arrValues(i) = Join(Array("TO_DATE('", values.Item(i).Value, "')"), "")
        End If

      '�o�C�i���^�E���[�W�I�u�W�F�N�g�^
      ElseIf getFormatGroupNameOrcl(Left(types.Item(i).Value, lenTimeStamp)) = "RAW" Then
        'DECODE('�f�[�^', 'HEX')
        arrValues(i) = Join(Array("HEXTORAW('", values.Item(i).Value, "')"), "")

      '��L�ɓ��Ă͂܂�Ȃ��f�[�^�^
      Else
        arrValues(i) = CStr(values.Item(i).Value)

      End If
    End If
  Next

  'INSERT���𐶐�
  INSERT_ORCL = Join(Array("INSERT INTO ", tableName, "(", Join(arrClmns, ","), ") VALUES(", Join(arrValues, ","), ");"), "")

End Function
