// Biểu giá mặc định (sinh từ Service_List_TanThuan_Vi.xlsx / _Eh.xlsx).
// Chỉ dùng khi người dùng chưa chỉnh sửa/nhập biểu giá riêng (lưu trong localStorage).
const DEFAULT_PRODUCTS = {
  "vi": [
    {
      "name": "Nâng / Hạ container tại Bãi <=> Xe - Cont 20' hàng",
      "unit": "cont",
      "price": 409320
    },
    {
      "name": "Nâng / Hạ container tại Bãi <=> Xe - Cont 20' rỗng",
      "unit": "cont",
      "price": 291600
    },
    {
      "name": "Nâng / Hạ container tại Bãi <=> Xe - Cont 40' hàng",
      "unit": "cont",
      "price": 690120
    },
    {
      "name": "Nâng / Hạ container tại Bãi <=> Xe - Cont 40' rỗng",
      "unit": "cont",
      "price": 439560
    },
    {
      "name": "Nâng / Hạ container tại Bãi <=> Xe - Cont trên 40' hàng",
      "unit": "cont",
      "price": 891000
    },
    {
      "name": "Nâng / Hạ container tại Bãi <=> Xe - Cont trên 40' rỗng",
      "unit": "cont",
      "price": 534600
    },
    {
      "name": "Xếp / Dỡ container Bãi <=> Sà lan - Cont 20' hàng",
      "unit": "cont",
      "price": 691740
    },
    {
      "name": "Xếp / Dỡ container Bãi <=> Sà lan - Cont 20' rỗng",
      "unit": "cont",
      "price": 353160
    },
    {
      "name": "Xếp / Dỡ container Bãi <=> Sà lan - Cont 40' hàng",
      "unit": "cont",
      "price": 1015740
    },
    {
      "name": "Xếp / Dỡ container Bãi <=> Sà lan - Cont 40' rỗng",
      "unit": "cont",
      "price": 536400
    },
    {
      "name": "Xếp / Dỡ container Bãi <=> Sà lan - Cont trên 40' hàng",
      "unit": "cont",
      "price": 1522800
    },
    {
      "name": "Xếp / Dỡ container Bãi <=> Sà lan - Cont trên 40' rỗng",
      "unit": "cont",
      "price": 806760
    },
    {
      "name": "Đóng / Rút hàng từ Sà lan <=> Container - Cont 20'",
      "unit": "cont",
      "price": 1512000
    },
    {
      "name": "Đóng / Rút hàng từ Sà lan <=> Container - Cont 40'",
      "unit": "cont",
      "price": 2160000
    },
    {
      "name": "Đóng hàng Xe => Container tại Bãi - Cont 20'",
      "unit": "cont",
      "price": 972000
    },
    {
      "name": "Đóng hàng Xe => Container tại Bãi - Cont 40'",
      "unit": "cont",
      "price": 1404000
    },
    {
      "name": "Rút hàng từ Container => Xe - Cont 20'",
      "unit": "cont",
      "price": 1080000
    },
    {
      "name": "Rút hàng từ Container => Xe - Cont 40'",
      "unit": "cont",
      "price": 1512000
    },
    {
      "name": "Sang hàng Cont => Cont - Cont 20'",
      "unit": "cont",
      "price": 2052000
    },
    {
      "name": "Sang hàng Cont => Cont - Cont 40'",
      "unit": "cont",
      "price": 2916000
    },
    {
      "name": "Sang hàng Cont => Cont - Cont 20' có Xe xúc",
      "unit": "cont",
      "price": 2667600
    },
    {
      "name": "Sang hàng Cont => Cont - Cont 40' có Xe xúc",
      "unit": "cont",
      "price": 3790800
    },
    {
      "name": "Sang hàng Cont => Cont - Cont 20' bao Jumbo",
      "unit": "cont",
      "price": 3078000
    },
    {
      "name": "Sang hàng Cont => Cont - Cont 40' bao Jumbo",
      "unit": "cont",
      "price": 4374000
    },
    {
      "name": "Đóng / Rút hàng từ Container <=> Kho - Cont 20'",
      "unit": "cont",
      "price": 1404000
    },
    {
      "name": "Đóng / Rút hàng từ Container <=> Kho - Cont 40'",
      "unit": "cont",
      "price": 2052000
    },
    {
      "name": "Vận chuyển container nội bộ Cảng - Cont 20' hàng",
      "unit": "cont",
      "price": 162000
    },
    {
      "name": "Vận chuyển container nội bộ Cảng - Cont 20' rỗng",
      "unit": "cont",
      "price": 118800
    },
    {
      "name": "Vận chuyển container nội bộ Cảng - Cont 40' hàng",
      "unit": "cont",
      "price": 237600
    },
    {
      "name": "Vận chuyển container nội bộ Cảng - Cont 40' rỗng",
      "unit": "cont",
      "price": 183600
    },
    {
      "name": "Cân container - Cont 20' xuất",
      "unit": "cont",
      "price": 86400
    },
    {
      "name": "Cân container - Cont 20' nhập",
      "unit": "cont",
      "price": 129600
    },
    {
      "name": "Cân container - Cont 40' xuất",
      "unit": "cont",
      "price": 129600
    },
    {
      "name": "Cân container - Cont 40' nhập",
      "unit": "cont",
      "price": 194400
    },
    {
      "name": "Vệ sinh container - Cont 20' quét / rửa nước",
      "unit": "cont",
      "price": 172800
    },
    {
      "name": "Vệ sinh container - Cont 40' quét / rửa nước",
      "unit": "cont",
      "price": 270000
    },
    {
      "name": "Vệ sinh môi trường - Cont 20'",
      "unit": "cont",
      "price": 70200
    },
    {
      "name": "Vệ sinh môi trường - Cont 40'",
      "unit": "cont",
      "price": 108000
    },
    {
      "name": "Chạy điện cho container lạnh - Cont 20'",
      "unit": "giờ/cont",
      "price": 32400
    },
    {
      "name": "Chạy điện cho container lạnh - Cont 40'",
      "unit": "giờ/cont",
      "price": 41040
    },
    {
      "name": "Lưu bãi cont 20' hàng",
      "unit": "ngày/cont",
      "price": 24840
    },
    {
      "name": "Lưu bãi cont 40' hàng",
      "unit": "ngày/cont",
      "price": 50760
    },
    {
      "name": "Lưu bãi cont trên 40' hàng",
      "unit": "ngày/cont",
      "price": 75600
    },
    {
      "name": "Lưu bãi cont 20' rỗng",
      "unit": "ngày/cont",
      "price": 12420
    },
    {
      "name": "Lưu bãi cont 40' rỗng",
      "unit": "ngày/cont",
      "price": 24840
    },
    {
      "name": "Lưu bãi cont trên 40' rỗng",
      "unit": "ngày/cont",
      "price": 37800
    },
    {
      "name": "Lưu bãi cont 20' lạnh hàng",
      "unit": "ngày/cont",
      "price": 37260
    },
    {
      "name": "Lưu bãi cont 40' lạnh hàng",
      "unit": "ngày/cont",
      "price": 76140
    },
    {
      "name": "Lưu bãi cont trên 40' lạnh hàng",
      "unit": "ngày/cont",
      "price": 113400
    },
    {
      "name": "Lưu bãi cont 20' lạnh rỗng",
      "unit": "ngày/cont",
      "price": 18630
    },
    {
      "name": "Lưu bãi cont 40' lạnh rỗng",
      "unit": "ngày/cont",
      "price": 37260
    },
    {
      "name": "Lưu bãi cont trên 40' lạnh rỗng",
      "unit": "ngày/cont",
      "price": 56700
    }
  ],
  "en": [
    {
      "name": "Lift On / Lift Off container Yard <=> Truck - Laden 20'",
      "unit": "cont",
      "price": 409320
    },
    {
      "name": "Lift On / Lift Off container Yard <=> Truck - Empty 20'",
      "unit": "cont",
      "price": 291600
    },
    {
      "name": "Lift On / Lift Off container Yard <=> Truck - Laden 40'",
      "unit": "cont",
      "price": 690120
    },
    {
      "name": "Lift On / Lift Off container Yard <=> Truck - Empty 40'",
      "unit": "cont",
      "price": 439560
    },
    {
      "name": "Lift On / Lift Off container Yard <=> Truck - Laden Over 40'",
      "unit": "cont",
      "price": 891000
    },
    {
      "name": "Lift On / Lift Off container Yard <=> Truck - Empty Over 40'",
      "unit": "cont",
      "price": 534600
    },
    {
      "name": "Loading / Unloading container Yard <=> Barge - Laden 20'",
      "unit": "cont",
      "price": 691740
    },
    {
      "name": "Loading / Unloading container Yard <=> Barge - Empty 20'",
      "unit": "cont",
      "price": 353160
    },
    {
      "name": "Loading / Unloading container Yard <=> Barge - Laden 40'",
      "unit": "cont",
      "price": 1015740
    },
    {
      "name": "Loading / Unloading container Yard <=> Barge - Empty 40'",
      "unit": "cont",
      "price": 536400
    },
    {
      "name": "Loading / Unloading container Yard <=> Barge - Laden Over 40'",
      "unit": "cont",
      "price": 1522800
    },
    {
      "name": "Loading / Unloading container Yard <=> Barge - Empty Over 40'",
      "unit": "cont",
      "price": 806760
    },
    {
      "name": "Stuffing / Unstuffing from Barge <=> Container - 20'",
      "unit": "cont",
      "price": 1512000
    },
    {
      "name": "Stuffing / Unstuffing from Barge <=> Container - 40'",
      "unit": "cont",
      "price": 2160000
    },
    {
      "name": "Stuffing cargo Truck => Container at Yard - 20'",
      "unit": "cont",
      "price": 972000
    },
    {
      "name": "Stuffing cargo Truck => Container at Yard - 40'",
      "unit": "cont",
      "price": 1404000
    },
    {
      "name": "Unstuffing cargo Container => Truck at Yard - 20'",
      "unit": "cont",
      "price": 1080000
    },
    {
      "name": "Unstuffing cargo Container => Truck at Yard - 40'",
      "unit": "cont",
      "price": 1512000
    },
    {
      "name": "Stuffing / Unstuffing Cont => Cont - 20'",
      "unit": "cont",
      "price": 2052000
    },
    {
      "name": "Stuffing / Unstuffing Cont => Cont - 40'",
      "unit": "cont",
      "price": 2916000
    },
    {
      "name": "Stuffing / Unstuffing Cont => Cont - 20' (with Forklift)",
      "unit": "cont",
      "price": 2667600
    },
    {
      "name": "Stuffing / Unstuffing Cont => Cont - 40' (with Forklift)",
      "unit": "cont",
      "price": 3790800
    },
    {
      "name": "Stuffing / Unstuffing Cont => Cont - 20' (with Jumbo bags)",
      "unit": "cont",
      "price": 3078000
    },
    {
      "name": "Stuffing / Unstuffing Cont => Cont - 40' (with Jumbo bags)",
      "unit": "cont",
      "price": 4374000
    },
    {
      "name": "Stuffing / Unstuffing from Container <=> Warehouse - 20'",
      "unit": "cont",
      "price": 1404000
    },
    {
      "name": "Stuffing / Unstuffing from Container <=> Warehouse - 40'",
      "unit": "cont",
      "price": 2052000
    },
    {
      "name": "Internal Port Shifting - Laden 20'",
      "unit": "cont",
      "price": 162000
    },
    {
      "name": "Internal Port Shifting - Empty 20'",
      "unit": "cont",
      "price": 118800
    },
    {
      "name": "Internal Port Shifting - Laden 40'",
      "unit": "cont",
      "price": 237600
    },
    {
      "name": "Internal Port Shifting - Empty 40'",
      "unit": "cont",
      "price": 183600
    },
    {
      "name": "Container Weighing - 20' (Export)",
      "unit": "cont",
      "price": 86400
    },
    {
      "name": "Container Weighing - 20' (Import)",
      "unit": "cont",
      "price": 129600
    },
    {
      "name": "Container Weighing - 40' (Export)",
      "unit": "cont",
      "price": 129600
    },
    {
      "name": "Container Weighing - 40' (Import)",
      "unit": "cont",
      "price": 194400
    },
    {
      "name": "Container Cleaning - 20' Sweeping / Washing by Water",
      "unit": "cont",
      "price": 172800
    },
    {
      "name": "Container Cleaning - 40' Sweeping / Washing by Water",
      "unit": "cont",
      "price": 270000
    },
    {
      "name": "Environmental Cleaning Fee - 20'",
      "unit": "cont",
      "price": 70200
    },
    {
      "name": "Environmental Cleaning Fee - 40'",
      "unit": "cont",
      "price": 108000
    },
    {
      "name": "Power Supply for Reefer Container - 20'",
      "unit": "hrs/cont",
      "price": 32400
    },
    {
      "name": "Power Supply for Reefer Container - 40'",
      "unit": "hrs/cont",
      "price": 41040
    },
    {
      "name": "Storage at Yard - Laden 20'",
      "unit": "day/cont",
      "price": 24840
    },
    {
      "name": "Storage at Yard - Laden 40'",
      "unit": "day/cont",
      "price": 50760
    },
    {
      "name": "Storage at Yard - Laden Over 40'",
      "unit": "day/cont",
      "price": 75600
    },
    {
      "name": "Storage at Yard - Empty 20'",
      "unit": "day/cont",
      "price": 12420
    },
    {
      "name": "Storage at Yard - Empty 40'",
      "unit": "day/cont",
      "price": 24840
    },
    {
      "name": "Storage at Yard - Empty Over 40'",
      "unit": "day/cont",
      "price": 37800
    },
    {
      "name": "Storage at Yard - Reefer Laden 20'",
      "unit": "day/cont",
      "price": 37260
    },
    {
      "name": "Storage at Yard - Reefer Laden 40'",
      "unit": "day/cont",
      "price": 76140
    },
    {
      "name": "Storage at Yard - Reefer Laden Over 40'",
      "unit": "day/cont",
      "price": 113400
    },
    {
      "name": "Storage at Yard - Reefer Empty 20'",
      "unit": "day/cont",
      "price": 18630
    },
    {
      "name": "Storage at Yard - Reefer Empty 40'",
      "unit": "day/cont",
      "price": 37260
    },
    {
      "name": "Storage at Yard - Reefer Empty Over 40'",
      "unit": "day/cont",
      "price": 56700
    }
  ]
};
