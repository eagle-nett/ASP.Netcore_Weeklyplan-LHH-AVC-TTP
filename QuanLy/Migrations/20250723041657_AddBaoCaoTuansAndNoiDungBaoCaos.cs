using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace QuanLy.Migrations
{
    /// <inheritdoc />
    public partial class AddBaoCaoTuansAndNoiDungBaoCaos : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_BaoCaoTuans_AspNetUsers_BaoCaoChoId",
                table: "BaoCaoTuans");

            migrationBuilder.DropForeignKey(
                name: "FK_BaoCaoTuans_AspNetUsers_NguoiBaoCaoId",
                table: "BaoCaoTuans");

            migrationBuilder.DropColumn(
                name: "ChuaHoanThanh",
                table: "NoiDungBaoCaos");

            migrationBuilder.DropColumn(
                name: "DaHoanThanh",
                table: "NoiDungBaoCaos");

            migrationBuilder.AlterColumn<string>(
                name: "TrachNhiemChinh",
                table: "NoiDungBaoCaos",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "",
                oldClrType: typeof(string),
                oldType: "nvarchar(max)",
                oldNullable: true);

            migrationBuilder.AlterColumn<string>(
                name: "MucDoUuTien",
                table: "NoiDungBaoCaos",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "",
                oldClrType: typeof(string),
                oldType: "nvarchar(max)",
                oldNullable: true);

            migrationBuilder.AddColumn<string>(
                name: "TienDo",
                table: "NoiDungBaoCaos",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");

            migrationBuilder.AlterColumn<string>(
                name: "BaoCaoChoId",
                table: "BaoCaoTuans",
                type: "nvarchar(450)",
                nullable: false,
                defaultValue: "",
                oldClrType: typeof(string),
                oldType: "nvarchar(450)",
                oldNullable: true);

            migrationBuilder.AddForeignKey(
                name: "FK_BaoCaoTuans_AspNetUsers_BaoCaoChoId",
                table: "BaoCaoTuans",
                column: "BaoCaoChoId",
                principalTable: "AspNetUsers",
                principalColumn: "Id",
                onDelete: ReferentialAction.Restrict);

            migrationBuilder.AddForeignKey(
                name: "FK_BaoCaoTuans_AspNetUsers_NguoiBaoCaoId",
                table: "BaoCaoTuans",
                column: "NguoiBaoCaoId",
                principalTable: "AspNetUsers",
                principalColumn: "Id",
                onDelete: ReferentialAction.Restrict);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_BaoCaoTuans_AspNetUsers_BaoCaoChoId",
                table: "BaoCaoTuans");

            migrationBuilder.DropForeignKey(
                name: "FK_BaoCaoTuans_AspNetUsers_NguoiBaoCaoId",
                table: "BaoCaoTuans");

            migrationBuilder.DropColumn(
                name: "TienDo",
                table: "NoiDungBaoCaos");

            migrationBuilder.AlterColumn<string>(
                name: "TrachNhiemChinh",
                table: "NoiDungBaoCaos",
                type: "nvarchar(max)",
                nullable: true,
                oldClrType: typeof(string),
                oldType: "nvarchar(max)");

            migrationBuilder.AlterColumn<string>(
                name: "MucDoUuTien",
                table: "NoiDungBaoCaos",
                type: "nvarchar(max)",
                nullable: true,
                oldClrType: typeof(string),
                oldType: "nvarchar(max)");

            migrationBuilder.AddColumn<string>(
                name: "ChuaHoanThanh",
                table: "NoiDungBaoCaos",
                type: "nvarchar(max)",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "DaHoanThanh",
                table: "NoiDungBaoCaos",
                type: "nvarchar(max)",
                nullable: true);

            migrationBuilder.AlterColumn<string>(
                name: "BaoCaoChoId",
                table: "BaoCaoTuans",
                type: "nvarchar(450)",
                nullable: true,
                oldClrType: typeof(string),
                oldType: "nvarchar(450)");

            migrationBuilder.AddForeignKey(
                name: "FK_BaoCaoTuans_AspNetUsers_BaoCaoChoId",
                table: "BaoCaoTuans",
                column: "BaoCaoChoId",
                principalTable: "AspNetUsers",
                principalColumn: "Id");

            migrationBuilder.AddForeignKey(
                name: "FK_BaoCaoTuans_AspNetUsers_NguoiBaoCaoId",
                table: "BaoCaoTuans",
                column: "NguoiBaoCaoId",
                principalTable: "AspNetUsers",
                principalColumn: "Id",
                onDelete: ReferentialAction.Cascade);
        }
    }
}
