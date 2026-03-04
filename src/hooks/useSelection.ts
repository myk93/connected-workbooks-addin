import { useState, useCallback, useEffect, useRef } from "react";
import { SelectionData } from "../types/addin";
import { getSelectedRange } from "../services/officeService";

interface UseSelectionResult {
    selection: SelectionData | null;
    loading: boolean;
    error: string | null;
    refresh: () => Promise<void>;
}

export function useSelection(): UseSelectionResult {
    const [selection, setSelection] = useState<SelectionData | null>(null);
    const [loading, setLoading] = useState(false);
    const [error, setError] = useState<string | null>(null);
    const debounceRef = useRef<ReturnType<typeof setTimeout> | null>(null);

    const refresh = useCallback(async () => {
        setLoading(true);
        setError(null);
        try {
            const data = await getSelectedRange();
            setSelection(data);
        } catch (err) {
            setError(err instanceof Error ? err.message : String(err));
            setSelection(null);
        } finally {
            setLoading(false);
        }
    }, []);

    useEffect(() => {
        // Initial load
        refresh();

        // Auto-refresh whenever the Excel selection changes.
        // Debounced at 350ms so rapid arrow-key navigation doesn't flood Excel.run calls.
        const handleChange = () => {
            if (debounceRef.current) clearTimeout(debounceRef.current);
            debounceRef.current = setTimeout(() => refresh(), 350);
        };

        Office.context.document.addHandlerAsync(
            Office.EventType.DocumentSelectionChanged,
            handleChange,
            () => {} // ignore async registration result
        );

        return () => {
            if (debounceRef.current) clearTimeout(debounceRef.current);
        };
    // eslint-disable-next-line react-hooks/exhaustive-deps
    }, []);

    return { selection, loading, error, refresh };
}
