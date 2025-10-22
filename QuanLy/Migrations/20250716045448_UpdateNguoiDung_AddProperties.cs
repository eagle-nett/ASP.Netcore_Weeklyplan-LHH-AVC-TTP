using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace QuanLy.Migrations
{
    /// <inheritdoc />
    public partial class UpdateNguoiDung_AddProperties : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_AspNetUsers_PhongBans_MaPhongBan",
                table: "AspNetUsers");

            migrationBuilder.AlterColumn<string>(
                name: "MaPhongBan",
                table: "AspNetUsers",
                type: "nvarchar(450)",
                nullable: false,
                defaultValue: "",
                oldClrType: typeof(string),
                oldType: "nvarchar(450)",
                oldNullable: true);

            migrationBuilder.AddColumn<string>(
                name: "ChucDanh",
                table: "AspNetUsers",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddColumn<string>(
                name: "ChucVu",
                table: "AspNetUsers",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddColumn<string>(
                name: "DienThoai",
                table: "AspNetUsers",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddColumn<string>(
                name: "GhiChu",
                table: "AspNetUsers",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddColumn<string>(
                name: "MaNV",
                table: "AspNetUsers",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddColumn<string>(
                name: "TenCongTy",
                table: "AspNetUsers",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddForeignKey(
                name: "FK_AspNetUsers_PhongBans_MaPhongBan",
                table: "AspNetUsers",
                column: "MaPhongBan",
                principalTable: "PhongBans",
                principalColumn: "MaPhongBan",
                onDelete: ReferentialAction.Cascade);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropForeignKey(
                name: "FK_AspNetUsers_PhongBans_MaPhongBan",
                table: "AspNetUsers");

            migrationBuilder.DropColumn(
                name: "ChucDanh",
                table: "AspNetUsers");

            migrationBuilder.DropColumn(
                name: "ChucVu",
                table: "AspNetUsers");

            migrationBuilder.DropColumn(
                name: "DienThoai",
                table: "AspNetUsers");

            migrationBuilder.DropColumn(
                name: "GhiChu",
                table: "AspNetUsers");

            migrationBuilder.DropColumn(
                name: "MaNV",
                table: "AspNetUsers");

            migrationBuilder.DropColumn(
                name: "TenCongTy",
                table: "AspNetUsers");

            migrationBuilder.AlterColumn<string>(
                name: "MaPhongBan",
                table: "AspNetUsers",
                type: "nvarchar(450)",
                nullable: true,
                oldClrType: typeof(string),
                oldType: "nvarchar(450)");

            migrationBuilder.AddForeignKey(
                name: "FK_AspNetUsers_PhongBans_MaPhongBan",
                table: "AspNetUsers",
                column: "MaPhongBan",
                principalTable: "PhongBans",
                principalColumn: "MaPhongBan");
        }
    }
}
