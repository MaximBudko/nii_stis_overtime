require 'docx'
require 'tempfile'

class OverWorkController < ApplicationController
  def generate_overtime_report

  end

  def do_generate_ov
    user_ids = params[:user_ids] || []
    @global_user_ids = user_ids
    start_date_str = params[:start_date]
    end_date_str = params[:end_date]
    type_overwork = params[:option_select]

    @global_type_overwork = type_overwork.to_s

    if user_ids.blank?
      flash[:error] = "Пользователи не выбраны"
      redirect_to action: :generate_overtime_report and return 
    end

    if start_date_str.blank?
      flash[:error] = "Укажите дату начала"
      redirect_to action: :generate_overtime_report and return 
    end

    if end_date_str.blank?
      flash[:error] = "Укажите дату окончания"
      redirect_to action: :generate_overtime_report and return 
    end 

    start_date = Date.parse(start_date_str)
    end_date = Date.parse(end_date_str)

    date_range = (start_date..end_date).to_a
    @date_range_overtime = date_range

    custom_field = TimeEntryCustomField.find_by(name: 'Тип работ')

    entries = TimeEntry
      .includes(:issue)
      .joins(:custom_values)
      .where(user_id: user_ids)
      .where(spent_on: date_range)
      .where(custom_values: {
        custom_field_id: custom_field.id,
        value: 'Сверхурочная'
      })

    @overtime_issues = entries.map(&:issue).compact.uniq
    @users = User.where(id: user_ids)

    if @overtime_issues.blank?
      flash[:error] = "У выбранных пользователей нет задач с типом работ 'Сверхурочная' за выбранный период."
      redirect_to action: :generate_overtime_report and return
    end

    respond_to do |format|
      format.xlsx do
        response.headers['Content-Disposition'] = 'attachment; filename=Отчет.xlsx'
        render xlsx: 'do_generate_ov', template: 'over_work/do_generate_ov'
      end
    end
  end

  # def generate_ov_doc_rep
  #   user_ids = params[:user_ids] || []
  #   @global_user_ids = user_ids
  #   start_date_str = params[:start_date]
  #   end_date_str = params[:end_date]
  #   type_overwork = params[:option_select]
  
  #   @global_type_overwork = type_overwork.to_s
  #   output_users = []
  #   users_loop = User.find(@global_user_ids)
  #   output_group = ""
  #   users_loop.each do |user|
  #     job_group = user.groups
  #     job_group.each do |group|
  #       if group.name == "Испытательная лаборатория"
  #         output_group = "испытательной лаборатории"
  #       elsif group.name == "Конструкторско-технологический отдел"
  #         output_group = "конструкторско-технологического отдела"
  #       elsif group.name == "Сектор встроенного программного обеспечения"
  #         output_group = "сектора встроенного программного обеспечения"
  #       elsif group.name == "Сектор источников питания"
  #         output_group = "сектора источников питания"
  #       elsif group.name == "Сектор конструирования РЭА"
  #         output_group = "сектора конструирования РЭА"
  #       elsif group.name == "Сектор печатных плат"
  #         output_group = "сектора печатных плат"
  #       elsif group.name == "Тематический отдел №1"
  #         output_group = "тематического отдела №1"
  #       elsif group.name == "Тематический отдел №2"
  #         output_group = "тематического отдела №2"
  #       elsif group.name == "Управление по научно-технической работе"
  #         output_group = "управления по научно-технической работе"
  #       end
  
  #       custom_field_users_surname = user.custom_field_values.detect { |cf| cf.custom_field.name == 'surname' }
  #       custom_field_users_job_title = user.custom_field_values.detect { |cf| cf.custom_field.name == 'job_title' }
  
  #       output_users << "#{custom_field_users_job_title} #{output_group} #{user.firstname.to_s[0]}.#{custom_field_users_surname.to_s[0]}. #{user.lastname} (Работы по ...)"
  #     end
  #   end
  
  #   #users_text = output_users.join("\n")
  
  #   # Выбор шаблона в зависимости от типа переработки
  #   template_filename = type_overwork == "option_1" ? "temp_wknd.docx" : "temp_ov.docx"
  
  #   # Новый путь до шаблона — с Рабочего стола
  #   template_path = File.expand_path("~/Desktop/template/#{template_filename}")
  
  #   unless File.exist?(template_path)
  #     render plain: "Шаблон не найден: #{template_path}", status: 404
  #     return
  #   end
  
  #   doc = Docx::Document.open(template_path)

  #   doc.paragraphs.each do |p|
  #     p.each_text_run do |r|
  #       if r.text.include?('{{user_and_users_data}}')
  #         r.text = '' # чистим место
  #       end
  #     end
  #   end
    
  #   # Потом уже добавляем новые параграфы
  #   output_users.each_with_index do |user_text, index|
  #     new_paragraph = doc.paragraphs.first.insert_paragraph_after("#{index + 1}. #{user_text}")
  #   end
        
  
  #   tmp_file = Tempfile.new(['generated_', '.docx'])
  #   doc.save(tmp_file.path)
  #   tmp_file.close
  
  #   send_file tmp_file.path,
  #             filename: "Служебная_записка_#{Time.now.strftime('%d.%m.%Y')}.docx",
  #             type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  #             disposition: 'attachment',
  #             stream: false
  
  #   # Удаление временного файла
  #   Thread.new do
  #     sleep 5
  #     tmp_file.unlink if File.exist?(tmp_file.path)
  #   end
  # end
  
  private

  def load_form_data
    @users = User.all
  end
end
