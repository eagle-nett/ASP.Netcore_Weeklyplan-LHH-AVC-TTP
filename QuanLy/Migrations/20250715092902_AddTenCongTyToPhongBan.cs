using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace QuanLy.Migrations
{
    /// <inheritdoc />
    public partial class AddTenCongTyToPhongBan : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<string>(
                name: "TenCongTy",
                table: "PhongBans",
                type: "nvarchar(max)",
                nullable: false,
                defaultValue: "");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "TenCongTy",
                table: "PhongBans");
        }
    }
}
