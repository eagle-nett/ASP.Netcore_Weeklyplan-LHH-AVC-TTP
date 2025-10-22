using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace QuanLy.Migrations
{
    /// <inheritdoc />
    public partial class AddExtraFieldsToNoiDungBaoCao : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<string>(
                name: "HuongGiaiQuyet",
                table: "NoiDungBaoCaos",
                type: "nvarchar(max)",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "KetQuaDatDuoc",
                table: "NoiDungBaoCaos",
                type: "nvarchar(max)",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "LyDoChuaHoanThanh",
                table: "NoiDungBaoCaos",
                type: "nvarchar(max)",
                nullable: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "HuongGiaiQuyet",
                table: "NoiDungBaoCaos");

            migrationBuilder.DropColumn(
                name: "KetQuaDatDuoc",
                table: "NoiDungBaoCaos");

            migrationBuilder.DropColumn(
                name: "LyDoChuaHoanThanh",
                table: "NoiDungBaoCaos");
        }
    }
}
