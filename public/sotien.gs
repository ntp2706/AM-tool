/**
 * Chuyển đổi số tiền thành chữ tiếng Việt.
 *
 * @param {number} so Số tiền cần đọc.
 * @param {string} donVi Đơn vị tiền tệ (mặc định là "đồng").
 * @return {string} Chuỗi đọc số tiền bằng chữ.
 * @customfunction
 * @example =DOC_SO_TIEN(A1)
 */
function DOC_SO_TIEN(so, donVi = "đồng") {
  if (so === null || so === undefined) return "";
  if (typeof so !== "number" && typeof so !== "string") return "Dữ liệu không hợp lệ";
  if (typeof so === "string" && so.trim() === "") return "";

  so = Number(so);
  if (!Number.isFinite(so)) return "Dữ liệu không hợp lệ";

  so = Math.round(so);
  if (!Number.isSafeInteger(so)) return "Số tiền vượt quá giới hạn hỗ trợ";
  if (so === 0) return "Không " + donVi;

  let am = false;
  if (so < 0) {
    am = true;
    so = Math.abs(so);
  }

  const chuSo = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"];
  const hang = ["", "nghìn", "triệu", "tỷ", "nghìn tỷ", "triệu tỷ"];

  function docBaSo(baSo, docKhongTram) {
    const tram = Math.floor(baSo / 100);
    const chuc = Math.floor((baSo % 100) / 10);
    const donvi = baSo % 10;
    let ketQua = "";

    if (tram > 0 || docKhongTram) {
      ketQua += chuSo[tram] + " trăm";
    }

    if (chuc > 1) {
      ketQua += (ketQua ? " " : "") + chuSo[chuc] + " mươi";
      if (donvi === 1) {
        ketQua += " mốt";
      } else if (donvi === 4) {
        ketQua += " tư";
      } else if (donvi === 5) {
        ketQua += " lăm";
      } else if (donvi > 0) {
        ketQua += " " + chuSo[donvi];
      }
    } else if (chuc === 1) {
      ketQua += (ketQua ? " " : "") + "mười";
      if (donvi === 5) {
        ketQua += " lăm";
      } else if (donvi > 0) {
        ketQua += " " + chuSo[donvi];
      }
    } else if (donvi > 0) {
      if (ketQua) {
        ketQua += " lẻ " + chuSo[donvi];
      } else {
        ketQua += chuSo[donvi];
      }
    }

    return ketQua;
  }

  let ketQua = "";
  const nhomSo = [];

  while (so > 0) {
    nhomSo.push(so % 1000);
    so = Math.floor(so / 1000);
  }

  for (let j = nhomSo.length - 1; j >= 0; j--) {
    const baSo = nhomSo[j];
    if (baSo > 0) {
      const docKhongTram = j < nhomSo.length - 1;
      const docNhom = docBaSo(baSo, docKhongTram);
      ketQua += (ketQua ? " " : "") + docNhom + (hang[j] ? " " + hang[j] : "");
    }
  }

  ketQua = (am ? "Âm " : "") + ketQua.trim();
  // Viết hoa chữ cái đầu tiên
  ketQua = ketQua.charAt(0).toUpperCase() + ketQua.slice(1);

  return ketQua + (donVi ? " " + donVi : "");
}