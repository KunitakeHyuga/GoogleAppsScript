function createDASH() {
  // Define your task list ID

  // Task title
  const taskTitle = 'ふうさんDASH';

  // Get today's date and set the due time to 23:59:59
  const today = new Date();
  today.setHours(23, 59, 59, 999);  // Set time to 23:59:59

  // Format the due date in ISO string format
  const dueDate = today.toISOString();

  // Task resource
  const task = {
    title: taskTitle,
    due: dueDate
  };

  // Add task to the task list
  const newTask = Tasks.Tasks.insert(task, listId);
  Logger.log('Task created: %s', newTask.title);
  Logger.log('Task ID: %s', newTask.id);
}
