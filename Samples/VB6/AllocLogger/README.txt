

Requirements:
   idbg debugger activex control from idefense (download installer)
   mscomctl.ocx  from microsoft
   vb6 runtimes

What is it;

	This app uses the idbg debugger to set breakpoints on:
		 GlobalAlloc, GlobalFree, LocalFree, LocalAlloc

	It logs:
		Where the buffer was allocated from (ret addr)
		How big the buffer is
		The buffer base address
		upon free - this record is updated to include the data in at time of freeing
		upon free - where was xxFree called from (ret addr)

	You can view the data as a text report, or hex editor view.
        remember data is only available on free. The hmem list entry will turn blue on free

	todo: add support for VirtualAlloc/VirtualFree, and HeapAlloc/HeapRealloc/HeapFree

	You can launch a new process, attach to an existing one, and pause/resume the process
        at anytime.

        I do have a complete olly like UI i have coupled with iDbg library which i was thinking
        of including as a probing tool if need be. will see if necessary. There are tons more
        logging that idbg supports that i could build into this. so if you have a cool idea let me know.


