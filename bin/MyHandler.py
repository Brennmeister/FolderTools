from watchdog.events import PatternMatchingEventHandler
class MyHandler(PatternMatchingEventHandler):
    def process(self, event):
        """
        event.event_type
            'modified' | 'created' | 'moved' | 'deleted'
        event.is_directory
            True | False
        event.src_path
            path/to/observed/file
        """
        # the file will be processed there
        #print('{:} - {:}'.format(event.src_path, event.event_type))  # print now only for degug

    def set_action_on_create(self,funhandle):
        if not hasattr(self, 'action_on_create'):
            self.action_on_create=list()
        self.action_on_create.append(funhandle)

    def remove_action_on_create(self,funhandle):
        self.action_on_create.remove(funhandle)

    def set_action_on_delete(self,funhandle):
        if not hasattr(self, 'action_on_delete'):
            self.action_on_delete=list()
        self.action_on_delete.append(funhandle)

    def remove_action_on_delete(self, funhandle):
        self.action_on_delete.remove(funhandle)


    def on_modified(self, event):
        self.process(event)

    def on_created(self, event):
        if hasattr(self, 'action_on_create'):
            for a in self.action_on_create:
                a([event.src_path]) # Call Action only for file with action!
        self.process(event)

    def on_deleted(self, event):
        if hasattr(self, 'action_on_delete'):
            for a in self.action_on_delete:
                a() # Call Action
        self.process(event)

    def on_renamed(self, event):
        self.process(event)
