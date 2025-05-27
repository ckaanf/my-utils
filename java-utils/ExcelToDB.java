class TestEquipmentServiceTest {

    private static final String DB_URL = "";
    private static final String DB_USER = "";
    private static final String DB_PASSWORD = "!";

    @Test
    void create() {
    }

    @Test
    void initDbFromExcel() throws Exception {
        // 엑셀 파일 경로 설정 (로컬 파일 경로)
        String filePath = "src/test/resources/test_equipment_data.xlsx";

        List<Clazz> equipmentList = new ArrayList<>();

        try (InputStream is = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(is)) {

            Sheet sheet = workbook.getSheetAt(0); // 첫 번째 시트 사용

            for (int i = 1; i <= 430; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                // 각 셀에서 값 추출
                String managementNo = getStringValue(row.getCell(0));
                String equipmentType = getStringValue(row.getCell(1));
                String equipmentName = getStringValue(row.getCell(2));
                Long amount = getLongValue(row.getCell(3));
                BigDecimal equipmentPrice = getLongValue(row.getCell(4)) != null ?
                    BigDecimal.valueOf(getLongValue(row.getCell(4))) : null;
                LocalDate purchaseDate = null;
                LocalDateTime dateTime = getDateValue(row.getCell(5));
                if (dateTime != null) {
                    purchaseDate = dateTime.toLocalDate();
                }
                String usefulLife = getStringValue(row.getCell(6));
                String managementStatus = getStringValue(row.getCell(7));
                String remark = getDateAsString(row.getCell(8));
                String damdangFirst = getStringValue(row.getCell(9));
                String damdangSecond = getStringValue(row.getCell(10));
                String location = getStringValue(row.getCell(11));

                // CATestEquipment 객체 생성
                Clazz equipment = Clazz.builder()
                        .build();

                equipmentList.add(equipment);
            }
        }

        // JDBC로 직접 DB에 저장
        try (Connection conn = DriverManager.getConnection(DB_URL, DB_USER, DB_PASSWORD)) {
            conn.setAutoCommit(false);

            String sql = "";

            try (PreparedStatement pstmt = conn.prepareStatement(sql)) {
                for (CATestEquipment equipment : equipmentList) {
                    pstmt.setString(1, equipment.getManagementNo());
                    pstmt.setString(2, equipment.getEquipmentType());
                    pstmt.setString(3, equipment.getEquipmentName());
                    pstmt.setObject(4, equipment.getAmount());
                    pstmt.setBigDecimal(5, equipment.getEquipmentPrice());
                    pstmt.setObject(6, equipment.getPurchaseDate());
                    pstmt.setString(7, equipment.getUsefulLife());
                    pstmt.setString(8, equipment.getUseYn());
                    pstmt.setString(9, equipment.getManagementStatus());
                    pstmt.setString(10, equipment.getRemark());
                    pstmt.setString(11, equipment.getDamdangFirst());
                    pstmt.setString(12, equipment.getDamdangSecond());
                    pstmt.setString(13, equipment.getLocation());
                    pstmt.setString(14, equipment.getRegId());
                    pstmt.setObject(15, equipment.getRegDate());
                    pstmt.setString(16, equipment.getModId());
                    pstmt.setObject(17, equipment.getModDate());

                    pstmt.addBatch();
                }

                int[] results = pstmt.executeBatch();
                conn.commit();

                int totalInserted = 0;
                for (int count : results) {
                    totalInserted += count;
                }
                System.out.println("엑셀 파일에서 " + totalInserted + "개의 장비 정보를 DB에 성공적으로 저장했습니다.");
                assertEquals(equipmentList.size(), totalInserted);
            } catch (SQLException e) {
                conn.rollback();
                throw e;
            }
        }
    }

    private String getStringValue(Cell cell) {
        if (cell == null) return null;
        
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf((long)cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            default:
                return null;
        }
    }
    
    private Long getLongValue(Cell cell) {
        if (cell == null) return null;
        
        switch (cell.getCellType()) {
            case NUMERIC:
                return (long) cell.getNumericCellValue();
            case STRING:
                try {
                    return Long.parseLong(cell.getStringCellValue());
                } catch (NumberFormatException e) {
                    return null;
                }
            default:
                return null;
        }
    }
    
    private LocalDateTime getDateValue(Cell cell) {
        if (cell == null) return null;
        
        switch (cell.getCellType()) {
            case NUMERIC:
                if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                    return cell.getLocalDateTimeCellValue();
                }
                return null;
            case STRING:
                try {
                    // 여러 형식의 날짜 문자열 처리
                    String dateStr = cell.getStringCellValue();
                    if (dateStr.contains(":")) {
                        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
                        return LocalDateTime.parse(dateStr, formatter);
                    } else {
                        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");
                        return LocalDateTime.of(java.time.LocalDate.parse(dateStr, formatter), java.time.LocalTime.MIDNIGHT);
                    }
                } catch (Exception e) {
                    return null;
                }
            default:
                return null;
        }
    }

    private String getDateAsString(Cell cell) {
        if (cell == null) return null;

        try {
            if (cell.getCellType() == CellType.NUMERIC) {
                double numericValue = cell.getNumericCellValue();

                LocalDate date = LocalDate.of(1899, 12, 30).plusDays((long) numericValue);

                return date.format(DateTimeFormatter.ofPattern("yyyy-MM-dd"));
            }
            else if (cell.getCellType() == CellType.STRING) {
                return cell.getStringCellValue();
            }

            return null;
        } catch (Exception e) {
            return null;
        }
    }
}