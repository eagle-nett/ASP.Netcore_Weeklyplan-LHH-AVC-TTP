using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace QuanLy.Migrations
{
    /// <inheritdoc />
    public partial class AddBaoCaoTuan : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "BaoCaoTuans",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    NguoiBaoCaoId = table.Column<string>(type: "nvarchar(450)", nullable: false),
                    BaoCaoChoId = table.Column<string>(type: "nvarchar(450)", nullable: true),
                    Tuan = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    TuNgay = table.Column<DateTime>(type: "datetime2", nullable: false),
                    DenNgay = table.Column<DateTime>(type: "datetime2", nullable: false),
                    NgayTao = table.Column<DateTime>(type: "datetime2", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_BaoCaoTuans", x => x.Id);
                    table.ForeignKey(
                        name: "FK_BaoCaoTuans_AspNetUsers_BaoCaoChoId",
                        column: x => x.BaoCaoChoId,
                        principalTable: "AspNetUsers",
                        principalColumn: "Id");
                    table.ForeignKey(
                        name: "FK_BaoCaoTuans_AspNetUsers_NguoiBaoCaoId",
                        column: x => x.NguoiBaoCaoId,
                        principalTable: "AspNetUsers",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateTable(
                name: "NoiDungBaoCaos",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    BaoCaoTuanId = table.Column<int>(type: "int", nullable: false),
                    NoiDung = table.Column<string>(type: "nvarchar(max)", nullable: false),
                    NgayHoanThanh = table.Column<DateTime>(type: "datetime2", nullable: true),
                    TrachNhiemChinh = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    TrachNhiemHoTro = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    MucDoUuTien = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    DaHoanThanh = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    ChuaHoanThanh = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    GhiChu = table.Column<string>(type: "nvarchar(max)", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_NoiDungBaoCaos", x => x.Id);
                    table.ForeignKey(
                        name: "FK_NoiDungBaoCaos_BaoCaoTuans_BaoCaoTuanId",
                        column: x => x.BaoCaoTuanId,
                        principalTable: "BaoCaoTuans",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateIndex(
                name: "IX_BaoCaoTuans_BaoCaoChoId",
                table: "BaoCaoTuans",
                column: "BaoCaoChoId");

            migrationBuilder.CreateIndex(
                name: "IX_BaoCaoTuans_NguoiBaoCaoId",
                table: "BaoCaoTuans",
                column: "NguoiBaoCaoId");

            migrationBuilder.CreateIndex(
                name: "IX_NoiDungBaoCaos_BaoCaoTuanId",
                table: "NoiDungBaoCaos",
                column: "BaoCaoTuanId");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "NoiDungBaoCaos");

            migrationBuilder.DropTable(
                name: "BaoCaoTuans");
        }
    }
}
