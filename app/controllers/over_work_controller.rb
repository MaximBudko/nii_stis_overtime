require 'docx'
require 'tempfile'

class OverWorkController < ApplicationController
  def generate_overtime_report

  end

  def users_by_date
    start_date = Date.parse(params[:start_date]) rescue nil
    end_date = Date.parse(params[:end_date]) rescue nil
  
    return render json: [] unless start_date && end_date
  
    date_range = start_date..end_date
  
    # Найдём кастомное поле "Тип работ"
    work_type_field = TimeEntryCustomField.find_by(name: 'Тип работ')
    return render json: [] unless work_type_field
  
    # Найдём time_entries, у которых Тип работ = Сверхурочная и они попадают в диапазон
    time_entries = TimeEntry.joins(:custom_values)
      .where(spent_on: date_range)
      .where(custom_values: {
        custom_field_id: work_type_field.id,
        value: 'Сверхурочная'
      })
  
    user_ids = time_entries.select(:user_id).distinct.pluck(:user_id)

  
    users = User.where(id: user_ids)

    user_data = users.map do |u|
      entry = time_entries.find { |te| te.user_id == u.id }
      {
        id: u.id,
        name: "#{u.firstname} #{u.lastname}",
        first_date: entry&.spent_on&.to_s
      }
    end

    render json: user_data
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
        render xlsx: 'Отчет о сверхурочной работе', template: 'over_work/do_generate_ov'
      end
    end
  end


  # def do_dokladnaya
  #   user_ids = params[:user_ids] || []
  #   @global_user_ids = user_ids
  #   start_date_str = params[:start_date]
  #   end_date_str = params[:end_date]
  #   type_overwork_dokladnaya = params[:option_select]

  #   @global_type_overwork_dokladnaya = type_overwork_dokladnaya.to_s
    
  #   p @global_type_overwork_dokladnaya

  #   if user_ids.blank?
  #     flash[:error] = "Пользователи не выбраны"
  #     redirect_to action: :generate_overtime_report and return 
  #   end

  #   if start_date_str.blank?
  #     flash[:error] = "Укажите дату начала"
  #     redirect_to action: :generate_overtime_report and return 
  #   end

  #   if end_date_str.blank?
  #     flash[:error] = "Укажите дату окончания"
  #     redirect_to action: :generate_overtime_report and return 
  #   end 

  #   start_date = Date.parse(start_date_str)
  #   end_date = Date.parse(end_date_str)

  #   date_range = (start_date..end_date).to_a
  #   @date_range_overtime = date_range

  #   custom_field = TimeEntryCustomField.find_by(name: 'Тип работ')

  #   entries = TimeEntry
  #     .includes(:issue)
  #     .joins(:custom_values)
  #     .where(user_id: user_ids)
  #     .where(spent_on: date_range)
  #     .where(custom_values: {
  #       custom_field_id: custom_field.id,
  #       value: 'Сверхурочная'
  #     })

  #   @overtime_issues = entries.map(&:issue).compact.uniq
  #   @users = User.where(id: user_ids)

  #   if @overtime_issues.blank?
  #     flash[:error] = "У выбранных пользователей нет задач с типом работ 'Сверхурочная' за выбранный период."
  #     redirect_to action: :generate_overtime_report and return
  #   end

  #   respond_to do |format|
  #     format.xlsx do
  #       response.headers['Content-Disposition'] = 'attachment; filename=Отчет.xlsx'
  #       render xlsx: 'do_dokladnaya', template: 'over_work/do_dokladnaya'
  #     end
  #   end
  # end
  
  private

  def load_form_data
    @users = User.all
  end
end
