using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SagyouKousuuKanriImport
{
    public partial class Frm_Import : Form
    {
        private const string COL_WORKDATE = "B";
        private const string COL_STARTTIME = "D";
        private const string COL_ENDTIME = "E";
        private const string COL_WORKNAIYOU = "H";
        private const string COL_STAFFID = "I";
        private const string COL_PROJECTID = "J";
        private const string COL_WORKID = "K";

        public Frm_Import()
        {
            InitializeComponent();
        }

        /// <summary>
        ///  参照ボタン　クリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btn_Refer_Click(object sender, EventArgs e)
        {
            // ファイル選択ダイアログを開く
            var ofDialog = new OpenFileDialog();

            if (ofDialog.ShowDialog() == DialogResult.OK)
            {
                Txt_FilePath.Text = ofDialog.FileName;
            }

            ofDialog.Dispose();
        }

        /// <summary>
        /// 取込ボタン　クリックイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Btn_Import_Click(object sender, EventArgs e)
        {
            try
            {
                // 入力チェック
                if (string.IsNullOrEmpty(Txt_FilePath.Text))
                {
                    MessageBox.Show("ファイルが指定されていません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Excel読み込み
                var importList = ReadData(Txt_FilePath.Text);

                // 取込データのチェック
                if (IsErrorImportData(importList))
                {
                    return;
                }

                // 登録処理
                Entry(importList);

                MessageBox.Show(importList.Count.ToString() + "件登録が完了しました。", "インフォメーション", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 接続文字列の取得
        /// </summary>
        /// <returns></returns>
        private string GetConnectionString()
        {
            var builder = new SqlConnectionStringBuilder()
            {
                DataSource = @"localhost\MSSQL2019",
                IntegratedSecurity = false,
                UserID = "sa",
                Password = "sqladmin",
                InitialCatalog = "DDW01_SagyouKousuuKanri"
            };

            return builder.ToString();
        }

        /// <summary>
        /// ファイルの読み込み処理
        /// </summary>
        /// <param name="path"></param>
        private List<Interface.ImportData> ReadData(string path)
        {
            var result = new List<Interface.ImportData>();

            int rowNo = 1;
            using (var workbook = new ClosedXML.Excel.XLWorkbook(path))
            {
                ClosedXML.Excel.IXLWorksheet sheet = workbook.Worksheet("Sheet1");

                // 読み込み処理
                Boolean headerflag = true;
                foreach (var r in sheet.Rows())
                {
                    // カウントアップ
                    rowNo++;

                    // 1行目はヘッダなので読み飛ばす
                    if (headerflag)
                    {
                        headerflag = false;
                        continue;
                    }

                    // 最終行までいったら終わる
                    if (r == sheet.LastRowUsed())
                    {
                        break;
                    }

                    // 登録項目がすべて未入力の場合は読み飛ばす
                    if (string.IsNullOrEmpty(r.Cell(COL_WORKDATE).Value.ToString()) && string.IsNullOrEmpty(r.Cell(COL_STARTTIME).Value.ToString())
                        && string.IsNullOrEmpty(r.Cell(COL_ENDTIME).Value.ToString()) && string.IsNullOrEmpty(r.Cell(COL_STAFFID).CachedValue.ToString())
                        && string.IsNullOrEmpty(r.Cell(COL_WORKNAIYOU).Value.ToString()) && string.IsNullOrEmpty(r.Cell(COL_PROJECTID).CachedValue.ToString())
                        && string.IsNullOrEmpty(r.Cell(COL_WORKID).CachedValue.ToString()))
                    {
                        continue;
                    }

                    //登録項目のいずれかが未入力の場合はエラー
                    if (string.IsNullOrEmpty(r.Cell(COL_WORKDATE).Value.ToString()) || string.IsNullOrEmpty(r.Cell(COL_STARTTIME).Value.ToString())
                        || string.IsNullOrEmpty(r.Cell(COL_ENDTIME).Value.ToString()) || string.IsNullOrEmpty(r.Cell(COL_STAFFID).CachedValue.ToString())
                        || string.IsNullOrEmpty(r.Cell(COL_WORKNAIYOU).Value.ToString()) || string.IsNullOrEmpty(r.Cell(COL_PROJECTID).CachedValue.ToString())
                        || string.IsNullOrEmpty(r.Cell(COL_WORKID).CachedValue.ToString()))
                    {
                        MessageBox.Show(rowNo.ToString() + "未入力の項目があります。");
                        return null;
                    }

                    DateTime workDate = (DateTime)r.Cell(COL_WORKDATE).Value;
                    DateTime startTime = (DateTime)r.Cell(COL_STARTTIME).Value;
                    DateTime editedStartTime = new DateTime(workDate.Year, workDate.Month, workDate.Day, startTime.Hour, startTime.Minute, startTime.Second);
                    DateTime endTime = (DateTime)r.Cell(COL_ENDTIME).Value;
                    DateTime editedEndTime = new DateTime(workDate.Year, workDate.Month, workDate.Day, endTime.Hour, endTime.Minute, endTime.Second);

                    result.Add(new Interface.ImportData
                    {
                        StaffID = long.Parse(r.Cell(COL_STAFFID).CachedValue.ToString()),
                        ProjectID = long.Parse(r.Cell(COL_PROJECTID).CachedValue.ToString()),
                        WorkID = long.Parse(r.Cell(COL_WORKID).CachedValue.ToString()),
                        WorkDate = workDate,
                        StartTime = editedStartTime,
                        EndTime = editedEndTime,
                        WorkNaiyou = r.Cell(COL_WORKNAIYOU).Value.ToString()
                    });
                }
            }
            return result;
        }

        /// <summary>
        /// 取込データのチェック
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        private bool IsErrorImportData(List<Interface.ImportData> list)
        {
            // 作業時間の整合性チェック
            var timeError = list.FindAll(x => x.StartTime >= x.EndTime);
            if (timeError.Count > 1)
            {
                MessageBox.Show("開始時間＞＝終了時間のデータがあります。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return true;
            }

            // 作業時間の重複チェック（取込データ内）
            var workdates = list.Select(x => x.WorkDate).Distinct();
            foreach (var workdate in workdates)
            {
                // 1日分のデータを抽出
                var oneday = list.Where(x => x.WorkDate == workdate);

                // 時間の重複を判定
                foreach (var item in oneday)
                {
                    // 開始時間が他の作業時間内にあるパターン
                    var a = oneday.Where(x => x.StartTime < item.StartTime).Where(y => item.StartTime < y.EndTime);
                    if (a.Count() > 0)
                    {
                        MessageBox.Show("取込データ内に作業時間の重複があります。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return true;
                    }
                    // 終了時間が他の作業時間内にあるパターン
                    var b = oneday.Where(x => x.StartTime < item.EndTime).Where(y => item.EndTime < y.EndTime);
                    if (b.Count() > 0)
                    {
                        MessageBox.Show("取込データ内に作業時間の重複があります。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return true;
                    }
                    // 他の作業時間を含んでいるパターン
                    var c = oneday.Where(x => x.StartTime >= item.StartTime).Where(y => item.EndTime >= y.EndTime);
                    if (c.Count() > 1)
                    {
                        MessageBox.Show("取込データ内に作業時間の重複があります。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return true;
                    }
                }
            }

            // 作業時間の重複チェック（登録済みデータ）
            foreach (var item in list)
            {
                if (IsErrorTimeDup(item))
                {
                    MessageBox.Show("登録済みデータとの作業時間の重複があります。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// 登録処理
        /// </summary>
        /// <param name="list"></param>
        private void Entry(List<Interface.ImportData> list)
        {
            using (var connection = new SqlConnection(GetConnectionString()))
            {
                try
                {
                    // データベースの接続開始
                    connection.Open();
                    var transaction = connection.BeginTransaction();
                    var command = connection.CreateCommand();
                    command.Transaction = transaction;

                    try
                    {
                        foreach (var item in list)
                        {
                            // 1レコードを新規登録
                            InsertNewData(command, item);
                        }
                    }
                    catch
                    {
                        // ロールバック
                        transaction.Rollback();
                        throw;
                    }
                    finally
                    {
                        // コミット
                        transaction.Commit();
                    }
                }
                catch (Exception exception)
                {
                    MessageBox.Show(exception.Message);
                    throw;
                }
                finally
                {
                    // データベースの接続終了
                    connection.Close();
                }
                //try
                //{
                //    // データベースの接続開始
                //    connection.Open();

                //    using (var transaction = connection.BeginTransaction())
                //    using (var command = new SqlCommand() { Connection = connection, Transaction = transaction })
                //    {
                //        try
                //        {
                //            foreach (var item in list)
                //            {
                //                // 1レコードを新規登録
                //                InsertNewData(command, item);
                //            }
                //        }
                //        catch
                //        {
                //            // ロールバック
                //            transaction.Rollback();
                //            throw;
                //        }
                //        finally
                //        {
                //            // コミット
                //            transaction.Commit();
                //        }
                //    }
                //}
                //catch (Exception exception)
                //{
                //    MessageBox.Show(exception.Message);
                //    throw;
                //}
                //finally
                //{
                //    // データベースの接続終了
                //    connection.Close();
                //}
            }
        }

        /// <summary>
        /// 時間の重複をチェック
        /// </summary>
        /// <param name="item"></param>
        private bool IsErrorTimeDup(Interface.ImportData item)
        {
            var table = new DataTable();

            using (var connection = new SqlConnection(GetConnectionString()))
            using (var command = connection.CreateCommand())
            {
                try
                {
                    // データベースの接続開始
                    connection.Open();

                    // SQLの設定
                    var sb = new StringBuilder();
                    sb.Append("SELECT ");
                    sb.Append("  WorkKousuuJisseki.* ");
                    sb.Append("FROM ");
                    sb.Append("  WorkKousuuJisseki ");
                    sb.Append("WHERE ");
                    sb.Append("  WorkKousuuJisseki.StaffID = @StaffID ");
                    sb.Append("  AND WorkKousuuJisseki.WorkDate = @WorkDate ");
                    sb.Append("  AND ( ");
                    sb.Append("    ( ");
                    sb.Append("      (WorkKousuuJisseki.StartTime < @StartTime) AND (@StartTime < WorkKousuuJisseki.EndTime) ");
                    sb.Append("    ) ");
                    sb.Append("    OR ( ");
                    sb.Append("      (WorkKousuuJisseki.StartTime < @EndTime) AND (@EndTime < WorkKousuuJisseki.EndTime) ");
                    sb.Append("    ) ");
                    sb.Append("    OR ( ");
                    sb.Append("      (WorkKousuuJisseki.StartTime >= @StartTime) AND (WorkKousuuJisseki.EndTime <= @EndTime) ");
                    sb.Append("    ) ");
                    sb.Append("  ) ");
                    command.CommandText = sb.ToString();
                    command.Parameters.Add(new SqlParameter("@StaffID", item.StaffID));
                    command.Parameters.Add(new SqlParameter("@WorkDate", item.WorkDate));
                    command.Parameters.Add(new SqlParameter("@StartTime", item.StartTime));
                    command.Parameters.Add(new SqlParameter("@EndTime", item.EndTime));

                    // SQLの実行
                    var adapter = new SqlDataAdapter(command);
                    adapter.Fill(table);

                    if (table.Rows.Count > 1)
                    {
                        return true;
                    }
                    return false;
                }
                catch (Exception exception)
                {
                    Console.WriteLine(exception.Message);
                    throw;
                }
                finally
                {
                    // データベースの接続終了
                    connection.Close();
                }
            }
        }

        /// <summary>
        /// 作業工数を新規登録
        /// </summary>
        /// <param name="command"></param>
        /// <param name="item"></param>
        private void InsertNewData(SqlCommand command, Interface.ImportData item)
        {
            // 実行するSQL
            var sb = new StringBuilder();
            sb.Append("INSERT ");
            sb.Append("INTO dbo.WorkKousuuJisseki( ");
            sb.Append("  StaffID ");
            sb.Append("  , ProjectID ");
            sb.Append("  , WorkID ");
            sb.Append("  , WorkDate ");
            sb.Append("  , StartTime ");
            sb.Append("  , EndTime ");
            sb.Append("  , WorkNaiyou ");
            sb.Append("  , TourokuDate ");
            sb.Append("  , KoushinDate ");
            sb.Append(") ");
            sb.Append("VALUES ( ");
            sb.Append("  @StaffID ");
            sb.Append("  , @ProjectID ");
            sb.Append("  , @WorkID ");
            sb.Append("  , @WorkDate ");
            sb.Append("  , @StartTime ");
            sb.Append("  , @EndTime ");
            sb.Append("  , @WorkNaiyou ");
            sb.Append("  , CURRENT_TIMESTAMP ");
            sb.Append("  , NULL ");
            sb.Append(") ");

            command.CommandText = sb.ToString();
            command.Parameters.Add(new SqlParameter("@StaffID", item.StaffID));
            command.Parameters.Add(new SqlParameter("@ProjectID", item.ProjectID));
            command.Parameters.Add(new SqlParameter("@WorkID", item.WorkID));
            command.Parameters.Add(new SqlParameter("@WorkDate", item.WorkDate));
            command.Parameters.Add(new SqlParameter("@StartTime", item.StartTime));
            command.Parameters.Add(new SqlParameter("@EndTime", item.EndTime));
            command.Parameters.Add(new SqlParameter("@WorkNaiyou", item.WorkNaiyou));

            // SQLの実行
            command.ExecuteNonQuery();
        }
    }
}