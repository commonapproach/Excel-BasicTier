import {
  Button,
  Checkbox,
  Dialog,
  DialogActions,
  DialogBody,
  DialogContent,
  DialogSurface,
  DialogTitle,
  Spinner,
} from "@fluentui/react-components";
import { Warning24Regular } from "@fluentui/react-icons";
import React, { useEffect, useMemo, useState } from "react";
import { FormattedMessage, useIntl } from "react-intl";
import { useDialogContext } from "../context/DialogContext";
import { prepareExportData, PreparedExport } from "../export/export";
import { getExportableOrganizations, OrgExportInfo, UNDATED_YEAR } from "../export/exportScan";
import { downloadJSONLD } from "../utils/utils";

interface ExportDialogProps {
  isDialogOpen: boolean;
  setDialogOpen: (isOpen: boolean) => void;
}

type Step = "organizations" | "years" | "summary";

const ExportDialog: React.FC<ExportDialogProps> = ({ isDialogOpen, setDialogOpen }) => {
  const { showDialog } = useDialogContext();
  const intl = useIntl();

  const [step, setStep] = useState<Step>("organizations");
  const [isScanning, setIsScanning] = useState(false);
  const [orgs, setOrgs] = useState<OrgExportInfo[]>([]);
  const [selectedOrgIds, setSelectedOrgIds] = useState<string[]>([]);
  const [selectedYears, setSelectedYears] = useState<string[]>([]);
  const [isPreparing, setIsPreparing] = useState(false);
  const [prepared, setPrepared] = useState<PreparedExport | null>(null);

  // Reset and scan the workbook whenever the dialog is opened.
  useEffect(() => {
    if (!isDialogOpen) return;

    let cancelled = false;
    setStep("organizations");
    setSelectedOrgIds([]);
    setSelectedYears([]);
    setOrgs([]);
    setPrepared(null);
    setIsPreparing(false);
    setIsScanning(true);

    (async () => {
      try {
        const result = await getExportableOrganizations();
        if (!cancelled) setOrgs(result);
      } catch (error: any) {
        if (!cancelled) {
          setDialogOpen(false);
          showDialog(
            `${intl.formatMessage({ id: "generics.error" })}!`,
            error?.message || intl.formatMessage({ id: "generics.error.message" }),
          );
        }
      } finally {
        if (!cancelled) setIsScanning(false);
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [isDialogOpen]);

  const selectedOrgs = useMemo(() => orgs.filter((o) => selectedOrgIds.includes(o.id)), [orgs, selectedOrgIds]);

  // Union of years across the selected organizations, newest first ("Undated" last).
  const availableYears = useMemo(() => {
    const years = new Set<string>();
    for (const org of selectedOrgs) {
      for (const year of Object.keys(org.reportsByYear)) years.add(year);
    }
    return Array.from(years).sort((a, b) => {
      if (a === UNDATED_YEAR) return 1;
      if (b === UNDATED_YEAR) return -1;
      return Number(b) - Number(a);
    });
  }, [selectedOrgs]);

  // Orgs that will actually be exported: a selected org is only included if it
  // has at least one indicator report in a selected year (others are dropped).
  const exportedOrgs = useMemo(
    () =>
      selectedOrgs.filter((o) =>
        Object.entries(o.reportsByYear).some(([year, count]) => count > 0 && selectedYears.includes(year)),
      ),
    [selectedOrgs, selectedYears],
  );

  const reportCountForSelection = useMemo(() => {
    let total = 0;
    for (const org of selectedOrgs) {
      for (const [year, count] of Object.entries(org.reportsByYear)) {
        if (selectedYears.includes(year)) total += count;
      }
    }
    return total;
  }, [selectedOrgs, selectedYears]);

  const toggleOrg = (id: string, checked: boolean) => {
    setSelectedOrgIds((prev) => (checked ? [...prev, id] : prev.filter((o) => o !== id)));
  };

  const allOrgsSelected = orgs.length > 0 && orgs.every((o) => selectedOrgIds.includes(o.id));

  const toggleAllOrgs = (checked: boolean) => {
    setSelectedOrgIds(checked ? orgs.map((o) => o.id) : []);
  };

  const toggleYear = (year: string, checked: boolean) => {
    setSelectedYears((prev) => (checked ? [...prev, year] : prev.filter((y) => y !== year)));
  };

  const allYearsSelected = availableYears.length > 0 && availableYears.every((y) => selectedYears.includes(y));

  const toggleAllYears = (checked: boolean) => {
    setSelectedYears(checked ? [...availableYears] : []);
  };

  const buildFileName = (): string => {
    if (exportedOrgs.length === 1) {
      const cleaned = exportedOrgs[0].name.replace(/[^\w]/gi, "");
      return cleaned || "Organization";
    }
    return "Multiple";
  };

  // Runs the export pipeline (build + filter + validate) without downloading,
  // so the summary screen can show warnings/errors before the user commits.
  const prepare = async () => {
    setIsPreparing(true);
    setPrepared(null);
    try {
      const result = await prepareExportData(intl, buildFileName(), {
        orgIds: selectedOrgIds,
        years: selectedYears,
      });
      setPrepared(result);
    } catch (error: any) {
      setDialogOpen(false);
      showDialog(
        `${intl.formatMessage({ id: "generics.error" })}!`,
        error?.message ||
          intl.formatMessage({
            id: "generics.error.message",
            defaultMessage: "Something went wrong",
          }),
      );
    } finally {
      setIsPreparing(false);
    }
  };

  const handleExport = () => {
    if (!prepared || prepared.errors.length > 0) return;
    downloadJSONLD(prepared.finalData, `${prepared.fileName}.json`);
    // Close the export wizard and show the success message after a short delay
    // so the browser's "Save As" dialog has time to appear first.
    setDialogOpen(false);
    setTimeout(() => {
      showDialog(
        intl.formatMessage({ id: "generics.success", defaultMessage: "Success" }),
        intl.formatMessage({
          id: "export.messages.success",
          defaultMessage: "Your data has been successfully exported.",
        }),
      );
    }, 1500);
  };

  const goNext = () => {
    if (step === "organizations") {
      setStep("years");
    } else if (step === "years") {
      setStep("summary");
      void prepare();
    }
  };

  const goBack = () => {
    if (step === "years") setStep("organizations");
    else if (step === "summary") setStep("years");
  };

  const captionStyle: React.CSSProperties = {
    color: "#2D62D7",
    fontStyle: "italic",
    fontSize: "0.8rem",
    marginTop: "0.5rem",
  };

  const yearsLabel = (years: string[]): string =>
    years
      .map((y) =>
        y === UNDATED_YEAR ? intl.formatMessage({ id: "export.year.undated", defaultMessage: "Undated" }) : y,
      )
      .join(", ");

  const renderOrganizationsStep = () => (
    <>
      <p>
        <FormattedMessage
          id="export.step.selectOrgs.description"
          defaultMessage="Select the organization(s) to export. Only organizations with at least one indicator report are listed."
        />
      </p>
      {orgs.length === 0 ? (
        <p style={{ fontWeight: "bold" }}>
          <FormattedMessage
            id="export.step.selectOrgs.empty"
            defaultMessage="No organizations with indicator reports were found. Add at least one indicator report before exporting."
          />
        </p>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", maxHeight: 280, overflowY: "auto" }}>
          <Checkbox
            checked={allOrgsSelected}
            onChange={(_, data) => toggleAllOrgs(!!data.checked)}
            label={intl.formatMessage({
              id: "export.step.selectOrgs.all",
              defaultMessage: "Select all organizations",
            })}
          />
          {orgs.map((org) => {
            const reportCount = Object.values(org.reportsByYear).reduce((a, b) => a + b, 0);
            return (
              <Checkbox
                key={org.id}
                checked={selectedOrgIds.includes(org.id)}
                onChange={(_, data) => toggleOrg(org.id, !!data.checked)}
                label={intl.formatMessage(
                  {
                    id: "export.step.selectOrgs.orgLabel",
                    defaultMessage: "{name} ({count} report(s))",
                  },
                  { name: org.name, count: reportCount },
                )}
              />
            );
          })}
        </div>
      )}
      {orgs.length > 0 && (
        <p style={captionStyle}>
          <FormattedMessage
            id="export.step.selectOrgs.caption"
            defaultMessage="Only data linked to the {count} selected organisation(s) will be exported."
            values={{ count: selectedOrgIds.length }}
          />
        </p>
      )}
    </>
  );

  const renderYearsStep = () => (
    <>
      <p>
        <FormattedMessage
          id="export.step.selectYears.description"
          defaultMessage="Select the reporting year(s) to include. Indicator reports outside these years will be excluded."
        />
      </p>
      <div style={{ display: "flex", flexDirection: "column", maxHeight: 280, overflowY: "auto" }}>
        <Checkbox
          checked={allYearsSelected}
          onChange={(_, data) => toggleAllYears(!!data.checked)}
          label={intl.formatMessage({
            id: "export.step.selectYears.all",
            defaultMessage: "All years",
          })}
        />
        {availableYears.map((year) => (
          <Checkbox
            key={year}
            checked={selectedYears.includes(year)}
            onChange={(_, data) => toggleYear(year, !!data.checked)}
            label={
              year === UNDATED_YEAR
                ? intl.formatMessage({ id: "export.year.undated", defaultMessage: "Undated" })
                : year
            }
          />
        ))}
      </div>
      {selectedYears.length > 0 && (
        <p style={captionStyle}>
          <FormattedMessage
            id="export.step.selectYears.caption"
            defaultMessage="Indicator reports for {years} selected."
            values={{ years: yearsLabel(selectedYears) }}
          />
        </p>
      )}
    </>
  );

  const renderSummaryStep = () => {
    if (isPreparing) {
      return (
        <div style={{ display: "flex", justifyContent: "center", padding: "1.5rem" }}>
          <Spinner
            label={intl.formatMessage({
              id: "export.preparing",
              defaultMessage: "Validating data...",
            })}
          />
        </div>
      );
    }

    const hasErrors = !!prepared && prepared.errors.length > 0;
    const hasWarnings = !!prepared && prepared.errors.length === 0 && prepared.warnings.length > 0;

    return (
      <div style={{ maxHeight: 320, overflowY: "auto" }}>
        <p>
          <FormattedMessage id="export.step.summary.intro" defaultMessage="The following data will be exported:" />
        </p>
        <ul style={{ display: "flex", flexDirection: "column", gap: "0.25rem", marginTop: "0.5rem" }}>
          <li>
            <FormattedMessage
              id="export.summary.orgs"
              defaultMessage="<b>{count} organization(s):</b> {names}"
              values={{
                b: (chunks: any) => <strong>{chunks}</strong>,
                count: exportedOrgs.length,
                names: exportedOrgs.map((o) => o.name).join(", "),
              }}
            />
          </li>
          <li>
            <FormattedMessage
              id="export.summary.years"
              defaultMessage="<b>Years:</b> {years}"
              values={{
                b: (chunks: any) => <strong>{chunks}</strong>,
                years: yearsLabel(selectedYears),
              }}
            />
          </li>
          <li>
            <FormattedMessage
              id="export.summary.reports"
              defaultMessage="<b>{count} indicator report(s)</b>"
              values={{
                b: (chunks: any) => <strong>{chunks}</strong>,
                count: reportCountForSelection,
              }}
            />
          </li>
        </ul>

        {hasErrors && (
          <div style={{ marginTop: "0.75rem" }}>
            <p style={{ fontWeight: "bold", color: "#b10e1c" }}>
              <FormattedMessage
                id="export.summary.errorsIntro"
                defaultMessage="Export is blocked until these issues are fixed:"
              />
            </p>
            <div dangerouslySetInnerHTML={{ __html: prepared!.errors.map((e) => `<p>${e}</p>`).join("") }} />
          </div>
        )}

        {hasWarnings && (
          <div style={{ marginTop: "0.75rem" }}>
            <p
              style={{
                fontWeight: "bold",
                display: "flex",
                alignItems: "center",
                gap: "0.35rem",
                color: "#B8860B",
              }}
            >
              <Warning24Regular />
              <FormattedMessage id="export.summary.warningsIntro" defaultMessage="Warnings:" />
            </p>
            <div dangerouslySetInnerHTML={{ __html: prepared!.warnings.join("<hr/>") }} />
          </div>
        )}
      </div>
    );
  };

  const stepTitleId =
    step === "organizations"
      ? "export.step.selectOrgs.title"
      : step === "years"
        ? "export.step.selectYears.title"
        : "export.step.summary.title";
  const stepTitleDefault =
    step === "organizations" ? "Select Organizations" : step === "years" ? "Select Years" : "Review & Export";

  const nextDisabled =
    (step === "organizations" && selectedOrgIds.length === 0) || (step === "years" && selectedYears.length === 0);

  // When the prepared export carries warnings, the user must explicitly choose
  // to export anyway, so the button reflects that.
  const exportHasWarnings = !!prepared && prepared.errors.length === 0 && prepared.warnings.length > 0;
  const exportLabelId = exportHasWarnings ? "export.button.ignoreAndExport" : "export.button.export";
  const exportLabelDefault = exportHasWarnings ? "Export Anyway" : "Export";

  return (
    <>
      {isDialogOpen && (
        <Dialog
          open={isDialogOpen}
          onOpenChange={(_, data) => {
            if (!data.open) setDialogOpen(false);
          }}
        >
          <DialogSurface>
            <DialogBody>
              <DialogTitle>
                <FormattedMessage id={stepTitleId} defaultMessage={stepTitleDefault} />
              </DialogTitle>
              <DialogContent>
                {isScanning ? (
                  <div style={{ display: "flex", justifyContent: "center", padding: "1.5rem" }}>
                    <Spinner
                      label={intl.formatMessage({
                        id: "export.loading",
                        defaultMessage: "Loading organizations...",
                      })}
                    />
                  </div>
                ) : step === "organizations" ? (
                  renderOrganizationsStep()
                ) : step === "years" ? (
                  renderYearsStep()
                ) : (
                  renderSummaryStep()
                )}
              </DialogContent>

              <DialogActions
                style={{
                  display: "flex",
                  flexDirection: "row",
                  justifyContent: "flex-end",
                  gap: "1rem",
                }}
              >
                <Button appearance="secondary" onClick={() => setDialogOpen(false)}>
                  <FormattedMessage id="generics.button.close" defaultMessage="Cancel" />
                </Button>
                {step !== "organizations" && (
                  <Button appearance="secondary" onClick={goBack} disabled={isScanning || isPreparing}>
                    <FormattedMessage id="export.button.back" defaultMessage="Back" />
                  </Button>
                )}
                {step === "summary" ? (
                  <Button
                    appearance="primary"
                    onClick={handleExport}
                    disabled={isScanning || isPreparing || !prepared || prepared.errors.length > 0}
                  >
                    <FormattedMessage id={exportLabelId} defaultMessage={exportLabelDefault} />
                  </Button>
                ) : (
                  <Button appearance="primary" onClick={goNext} disabled={isScanning || nextDisabled}>
                    <FormattedMessage id="generics.button.next" defaultMessage="Next" />
                  </Button>
                )}
              </DialogActions>
            </DialogBody>
          </DialogSurface>
        </Dialog>
      )}
    </>
  );
};

export default ExportDialog;
