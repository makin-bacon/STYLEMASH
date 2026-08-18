interface HelpModalProps {
  onClose: () => void
}

/** Placeholder help modal, opened from AppHeader's "Help" button. Content is
 * filler copy - swap in real walkthrough/FAQ text when it's written. */
export function HelpModal({ onClose }: HelpModalProps) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 px-4">
      <div className="w-full max-w-md rounded-lg bg-white p-5 shadow-xl">
        <h2 className="text-3xl font-semibold text-slate-800">Help</h2>
        <div className="mt-3 max-h-80 space-y-3 overflow-y-auto text-sm text-slate-600">
            <h3 className="text-2xl font-semibold text-slate-800">Welcome to the StyleMash.</h3>
          <p>
            This is meant to be a simple tool for rapidly cleaning messy Word files. Just upload your Word file and StyleMash will analyse the styles in the document. Once analysed you can start merging your duplicate and unnamed styles into approved styles.
          </p>
          <h4 className="text-lg font-semibold text-slate-800">Wait, what styles?</h4>
          <p>
            StyleMash needs a bit of help to clean your documents. Either create new styles from scratch in the right hand panel after loading you source file, or, even quicker, upload a new Word file that has a list of known, clean styles and then use those as targets for merging.
          </p>
          <h4 className="text-lg font-semibold text-slate-800">How do I merge these things?</h4>
          <p>
            Merging is easy. First, either create styles by clicking the "New styles" button or upload a Word file of your own to populate the list on the right. Select any new or non-approved styles from the list on the left - you can select as many or as few as you like - and then selct the style on the right you want to merge it with and then click the "Do it" button. 
          </p>
          <h4 className="text-lg font-semibold text-slate-800">How much to merge?</h4>
          <p>
            Ideally, your documents should only use a limited number of approved styles for consistency across channels and to aid in production. For a copy of an approved style source file for use in StyleMash, contact <a href="mailto:design@scu.edu.au">design@scu.edu.au</a>. 
          </p>
          <h4 className="text-lg font-semibold text-slate-800">Can I use this on my phone?</h4>
          <p>
            The short answer is no. There are no plans to make this a mobile-friendly release. The long answer involves talking about using MS Word on your phone and I'd rather not. 
          </p>
        </div>
        <div className="mt-4 flex justify-end">
          <button
            type="button"
            onClick={onClose}
            className="rounded-md bg-indigo-600 px-3 py-1.5 text-xs font-medium text-white hover:bg-indigo-700"
          >
            Close
          </button>
        </div>
      </div>
    </div>
  )
}
