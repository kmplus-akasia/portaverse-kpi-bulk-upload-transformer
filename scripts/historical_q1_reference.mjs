export function classifyAssignment(row) {
  if (row.lakhar_id != null) return "LAKHAR";
  if (row.job_sharing_id != null) return "JOB_SHARING";
  return "PRIMARY";
}

export function buildHistoricalAssignmentQuery() {
  return `
    SELECT tepms.employee_position_master_sync_id,
           tepms.employee_number,
           tepms.position_master_variant_id,
           tepms.start_date AS assignment_start_date,
           tepms.end_date AS assignment_end_date,
           tepms.lakhar_id,
           tepms.job_sharing_id,
           TRIM(CONCAT_WS(' ', te.firstname, te.middlename, te.lastname)) AS employee_name,
           tpmv.position_master_id,
           tpmv.position_master_variant_id AS variant_id,
           tpm.name AS position_title,
           tpm.position_master_type_id,
           tpm.job_class_level,
           tpmos.position_master_organization_sync_id,
           tpmos.organization_master_id AS group_master_id,
           tpmos.start_date AS organization_start_date,
           tpmos.end_date AS organization_end_date,
           tgm.name AS group_name,
           tgm.company_id,
           tgm.org_level AS group_org_level,
           tgm.org_type AS group_org_type,
           tgm.costcenter AS group_costcenter,
           tci.name AS company_name,
           tci.code AS company_code,
           CASE
             WHEN tpmos.position_master_organization_sync_id IS NULL
               OR tgm.group_master_id IS NULL
               OR tci.company_in_id IS NULL
             THEN 1 ELSE 0
           END AS missing_historical_organization
      FROM tb_employee_position_master_sync tepms
      LEFT JOIN tb_employee te
        ON te.employee_number = tepms.employee_number
      LEFT JOIN tb_position_master_variant tpmv
        ON tpmv.position_master_variant_id = tepms.position_master_variant_id
      LEFT JOIN tb_position_master_v2 tpm
        ON tpm.position_master_id = tpmv.position_master_id
      LEFT JOIN tb_position_master_organization_sync tpmos
        ON tpmos.position_master_id = tpm.position_master_id
       AND tpmos.deletedAt IS NULL
       AND DATE(tepms.end_date) BETWEEN COALESCE(tpmos.start_date, '1000-01-01')
                                   AND COALESCE(tpmos.end_date, '9999-12-31')
      LEFT JOIN tb_group_master tgm
        ON tgm.group_master_id = tpmos.organization_master_id
       AND tgm.deletedAt IS NULL
       AND DATE(tepms.end_date) BETWEEN COALESCE(tgm.start_date, '1000-01-01')
                                   AND COALESCE(tgm.end_date, '9999-12-31')
      LEFT JOIN tb_company_in tci
        ON tci.company_in_id = tgm.company_id
       AND tci.deletedAt IS NULL
       AND DATE(tepms.end_date) BETWEEN COALESCE(tci.start_date, '1000-01-01')
                                   AND COALESCE(tci.end_date, '9999-12-31')
     WHERE tepms.deletedAt IS NULL
       AND DATE(tepms.end_date) = ?
       AND (tgm.company_id = ? OR tgm.group_master_id IS NULL)
     ORDER BY tepms.employee_number ASC, tepms.position_master_variant_id ASC,
              tpmos.organization_master_id ASC`;
}

export function buildHistoricalNomenclatureQuery() {
  return `
    SELECT pnm.id,
           pnm.cluster_id,
           pnm.cluster_label,
           pnm.position_master_id,
           pnm.job_class_level,
           pnm.position_name,
           pnm.position_master_type_id,
           pnm.type_name,
           pnm.group_master_id,
           pnm.group_name,
           pnm.company_id,
           pnm.company_name,
           tgm.name AS current_group_name,
           tci.name AS current_company_name,
           tci.code AS current_company_code
      FROM position_nomenclature_mapping pnm
      LEFT JOIN tb_group_master tgm
        ON tgm.group_master_id = pnm.group_master_id
      LEFT JOIN tb_company_in tci
        ON tci.company_in_id = tgm.company_id
     WHERE pnm.position_master_id IS NOT NULL
       AND pnm.company_id = ?
     ORDER BY pnm.position_master_id ASC, pnm.cluster_id ASC, pnm.group_master_id ASC`;
}

export function shapeHistoricalPayload({
  profile,
  cutoffDate,
  companyId,
  assignmentRows,
  nomenclatureRows,
}) {
  return {
    source: {
      profile,
      cutoff_date: cutoffDate,
      company_id: String(companyId),
      read_only: true,
    },
    historical_assignment_rows: assignmentRows.map((row) => ({
      ...row,
      assignment_type: classifyAssignment(row),
    })),
    nomenclature_rows: nomenclatureRows,
  };
}
